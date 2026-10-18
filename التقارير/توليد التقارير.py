# -*- coding: utf-8 -*-
"""Generate the 4 HTML audit reports with clickable source references."""
import openpyxl, json, re, os, sys
sys.path.insert(0, 'C:/Users/engmh/alrawda_check')
from refs_orig import MONEY_REFS, ACCT_REFS

DB = '../index.html'
MAN = json.load(open('C:/Users/engmh/OneDrive/Documents/ghazali check/قاعدة البيانات/data/manifest.json', encoding='utf-8'))
SID = {(x['src'], x['sheet']): x for x in MAN}
MF = json.load(open('C:/Users/engmh/alrawda_check/money_final.json', encoding='utf-8'))
MP = json.load(open('C:/Users/engmh/alrawda_check/money_report_parts.json', encoding='utf-8'))
NAT = 'C:/Users/engmh/OneDrive/Documents/ghazali check/نتيجة/'
OUT = 'C:/Users/engmh/OneDrive/Documents/ghazali check/قاعدة البيانات/التقارير/'
os.makedirs(OUT, exist_ok=True)

def xl_url(src, sheet, excel_row=None):
    m = SID.get((src, sheet))
    if not m: return None
    u = f"{DB}?s={m['id']}"
    if excel_row:
        r = excel_row - 3
        if 0 <= r < m['n']: u += f"&r={r}"
    return u

def link(url, text):
    if not url: return text
    return f'<a class="ref" href="{url}" target="_blank">{text}</a>'

def file_of(t):
    if 'الفترة الاولي' in t or 'الفترة الأولى' in t or 'الحركة ف1' in t: return 'P1'
    if 'حركه تعديل' in t: return 'PT'
    if 'الفترة الثانية' in t or '2الحركة' in t: return 'P2'
    if 'الخزينة' in t: return 'KH'
    if re.search(r'[\s(]ف1(?![0-9])', t): return 'P1'
    if re.search(r'[\s(]ف2(?![0-9])', t): return 'P2'
    return None

PAIR = {'مشتريات وموردين': 'مبيعات وعملاء', 'مبيعات وعملاء': 'مشتريات وموردين'}
ROW_AFTER = r'[^0-9:]{0,14}(?:الصفوف?|سطر|السطر|صف)\s+(\d[\d,]*)|:\s*(\d[\d,]*)\s*(?:إلى|·|,|\)|$)'

def parse_src_file(t):
    return file_of(t)

SHEETS = ['مشتريات وموردين','مبيعات وعملاء','الخزينة','مصروفات','مسير مرتبات','حساب المضرب',
          'ميزان عملاء','ميزان موردين','Sheet1','Sheet2','Sheet3','Sheet4','Sheet5','Sheet6','Sheet8',
          'ورقة1','ورقة2','ورقة3','الشونه','شعير من الشونة','كونة مخزن نجم','الغلة','الغلة 2025','برسيم']

def linkify(text, files_ctx='', sheets_ctx=''):
    """Convert recognizable source mentions into links. Returns HTML."""
    # program invoice refs: جدول sadr/ward + سطر فاتورة N
    def inv_repl(m):
        tbl, num = m.group(1), m.group(2)
        sid = 123 if 'sadr' in tbl else 124
        return f'<a class="ref" href="{DB}?s={sid}&q=فاتورة {num}" target="_blank">فاتورة {num}</a>'
    text = re.sub(r'جدول (sadr|ward)[^/]*?/سطر فاتورة (\d+)', inv_repl, text)
    def_sheet = next((s for s in SHEETS if s in sheets_ctx), None)
    # sheet mentions with optional row ref right after (reverse order: safe splicing)
    matches = list(re.finditer(r'(مشتريات وموردين|مبيعات وعملاء|الخزينة|مصروفات|مسير مرتبات|حساب المضرب|ميزان عملاء|ميزان موردين|Sheet\d|ورقة\d|الشونه|شعير من الشونة|كونة مخزن نجم|الغلة 2025|الغلة|برسيم)', text))
    for m in reversed(matches):
        sheet = m.group(1)
        after = text[m.end():m.end()+30]
        rm = re.match(ROW_AFTER, after)
        num = (rm.group(1) or rm.group(2)) if rm else None
        src = file_of(text[max(0,m.start()-150):m.start()]) or file_of(files_ctx)
        if num and src:
            u = xl_url(src, sheet, int(num.replace(',', '')))
            if u:
                cut = m.end() + rm.end(0)
                text = text[:m.start()] + sheet + ' ' + link(u, 'سطر '+num) + text[cut:]
        elif src:
            u = xl_url(src, sheet)
            if u:
                text = text[:m.start()] + link(u, sheet) + text[m.end():]
    # bare row refs (no sheet nearby): sheet from window hints or default
    if def_sheet:
        refs = list(re.finditer(r'(?:الصفوف?|سطر|السطر)\s+(\d[\d,]*)', text))
        for m in reversed(refs):
            if text[:m.start()].count('<a ') > text[:m.start()].count('</a>'):
                continue  # inside an existing link
            win = text[max(0,m.start()-50):m.start()]
            sh = next((s for s in SHEETS if s in win), None)
            if not sh:
                if '(بيع)' in win and def_sheet in PAIR: sh = PAIR[def_sheet]
                elif '(شراء)' in win and def_sheet in PAIR: sh = def_sheet if 'مشتريات' in def_sheet else PAIR.get(def_sheet, def_sheet)
                else: sh = def_sheet
            if '(شراء)' in win: sh = 'مشتريات وموردين'
            if '(بيع)' in win: sh = 'مبيعات وعملاء'
            src = file_of(win) or file_of(files_ctx)
            if src:
                u = xl_url(src, sh, int(m.group(1).replace(',', '')))
                if u:
                    text = text[:m.start()] + link(u, 'سطر '+m.group(1)) + text[m.end():]
    return text

def fmt(n):
    try: return f"{float(n):,.0f}"
    except: return str(n)

CSS = """<style>
@page{size:A4;margin:10mm}
*{box-sizing:border-box}
body{font-family:'Segoe UI',Tahoma,Arial,sans-serif;direction:rtl;margin:0;color:#1a1a1a;font-size:12px;line-height:1.6;background:#fff}
.wrap{max-width:1000px;margin:0 auto;padding:16px}
h1{font-size:19px;text-align:center;background:#1f4d2e;color:#fff;padding:12px;border-radius:8px;margin:0 0 6px}
.sub{text-align:center;color:#555;font-size:12px;margin-bottom:12px}
h2{font-size:15px;background:#e8f0e4;border-right:5px solid #1f4d2e;padding:6px 12px;margin:18px 0 8px}
table{width:100%;border-collapse:collapse;margin-bottom:8px}
th{background:#1f4d2e;color:#fff;padding:6px;font-size:12px}
td{border:1px solid #bbb;padding:5px 8px;font-size:12px;vertical-align:top}
tr:nth-child(even) td{background:#f2f5ee}
td.n,th.n{text-align:left;font-variant-numeric:tabular-nums;white-space:nowrap}
td.c,th.c{text-align:center}
tr.total td{background:#dff0d8;font-weight:bold}
tr.grand td{background:#1f4d2e;color:#fff;font-weight:bold;font-size:14px}
.cards{display:flex;gap:8px;margin:10px 0;flex-wrap:wrap}
.card{flex:1;min-width:140px;border:2px solid #1f4d2e;border-radius:8px;padding:10px;text-align:center}
.card .v{font-size:17px;font-weight:bold}
.card .l{font-size:12px;color:#444}
a.ref{background:#fff3c4;border:1px solid #c9a100;border-radius:5px;padding:0 7px;text-decoration:none;color:#5a4a00;font-size:11.5px;white-space:nowrap}
a.ref:hover{background:#ffe27a}
.refbox{background:#f7f9f4;border:1px solid #cfe0c8;border-radius:8px;padding:8px 12px;margin:6px 0}
.refbox b.code{background:#1f4d2e;color:#fff;border-radius:5px;padding:0 8px;font-size:12px}
.note{background:#fff8e1;border:1px solid #d4a017;border-radius:6px;padding:8px 12px;margin:8px 0}
.how{color:#333;font-size:12px}
.src{color:#666;font-size:11px}
@media print{a.ref{text-decoration:none}}
.tbl{overflow-x:auto;margin-bottom:8px;-webkit-overflow-scrolling:touch}
@media(max-width:640px){.tbl table{min-width:560px}body{font-size:13px}.wrap{padding:10px}h1{font-size:16px}.brand h1{font-size:18px}.cards{gap:6px}.card{min-width:46%;padding:8px}.card .v{font-size:15px}td,th{font-size:12px;padding:4px 6px}a.ref{font-size:12px}}
</style>"""

BRAND = """<div class="brand"><span class="star">★</span><div><h1>مضرب الدمياطي استار</h1><div class="tag">تقارير المراجعة والحسابات</div></div></div>"""

BCSS = """<style>
.brand{display:flex;align-items:center;gap:10px;background:#12351f;color:#fff;border-radius:10px;padding:10px 14px;margin-bottom:4px;border-bottom:3px solid #c9a100}
.brand .star{font-size:32px;color:#f2c400;line-height:1}
.brand h1{margin:0;font-size:19px}
.brand .tag{font-size:12px;color:#ddd}
</style>"""

def numfmt(c):
    s = str(c)
    try:
        f = float(s.replace(',', ''))
        if '.' not in s and abs(f) < 10000: return s
        if abs(f) >= 1000: return f"{f:,.0f}"
        return f"{f:,.1f}"
    except: return c

def page(title, sub, body):
    return f"""<html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>{title} - مضرب الدمياطي استار</title>{CSS}{BCSS}</head>
<body><div class="wrap">{BRAND}<h1>{title}</h1><div class="sub">{sub}</div>{body}
<div class="note">طريقة الاستخدام: أي زر أصفر <a class="ref">مرجع</a> لما تدوس عليه يفتح مكان المصدر في قاعدة البيانات على نفس السطر.</div>
</div></body></html>"""

# ================= R1 =================
def build_r1():
    sell, buy = MF['sell'], MF['buy']
    MV_BUY = 10535500.0 + 2742708.6; MV_SELL = 44000.0
    nS = sell['total'] - MV_SELL; nB = buy['total'] - MV_BUY; MOVED = MV_BUY + MV_SELL
    nE = MP['TOTAL_EXP'] + MOVED; nNET = (nS - nB) - nE
    exp_e1 = MF.get('مصروفات_اكسل_ف1', MF.get('exp_e1', {}))
    cards = f"""<div class="cards">
<div class="card"><div class="v">{fmt(nS)}</div><div class="l">إجمالي البيع بعد نقل النولون (ج)</div></div>
<div class="card"><div class="v">{fmt(nB)}</div><div class="l">إجمالي الشراء بعد نقل الأجرة (ج)</div></div>
<div class="card"><div class="v">{fmt(nS-nB)}</div><div class="l">الربح قبل المصروفات (ج)</div></div>
<div class="card"><div class="v">{fmt(nE)}</div><div class="l">إجمالي المصروفات بعد النقل (ج)</div></div>
<div class="card"><div class="v">{fmt(nNET)}</div><div class="l">صافي الربح بعد النقل (ج)</div></div>
<div class="card"><div class="v">{abs(sell['tons']-buy['tons']):,.1f}</div><div class="l">مخزون آخر المدة (طن)</div></div>
</div>"""
    exp_rows = [
        ('مصروفات الإكسل (مرة واحدة)', 844403, 'م6', xl_url('P1','مصروفات',17)),
        ('مرتبات وأجور العمال حتى 30/9/2025', 1949578, 'م7', xl_url('P2','مسير مرتبات')),
        ('مصروفات البرنامج 2/10/2025 إلى 22/8/2026', 4188879, 'م8', None),
        ('أجرة المضرب والناولون (منقولة من المشتريات والمبيعات)', MOVED, 'م8-أجرة', None),
    ]
    t = '<h2>المصروفات مجمعة</h2><table><tr><th>البند</th><th class="n">القيمة (ج)</th><th class="c">المرجع</th></tr>'
    for name, val, code, u in exp_rows:
        t += f'<tr><td>{name}</td><td class="n">{fmt(val)}</td><td class="c">{link(u, code)}</td></tr>'
    t += f'<tr class="grand"><td>الإجمالي بعد النقل</td><td class="n">{fmt(nE)}</td><td></td></tr></table>'
    t += '<h2>إعادة تبويب أجرة المضرب (مصروف وليست شراء)</h2><div class="note">أجرة المضرب مصروفات مش شراء، وكده كده هتتضاف تاني كربح لأن المضرب ملك الملاك.</div><table><tr><th>البند المنقول</th><th class="n">القيمة (ج)</th><th>المرجع</th></tr>'
    t += f"<tr><td>ضرب أرز (6 سندات ward)</td><td class='n'>{fmt(10535500.0)}</td><td class='src'>{linkify('جدول ward (مشتريات)/سطر فاتورة 1864 · سطر 2054')}</td></tr>"
    t += f"<tr><td>نولون مشتريات (947 سند ward)</td><td class='n'>{fmt(2742708.6)}</td><td class='src'>{linkify('جدول ward (مشتريات)/سطر فاتورة 1008 · سطر 1115')}</td></tr>"
    t += f"<tr><td>نولون مبيعات (7 سندات sadr)</td><td class='n'>{fmt(44000.0)}</td><td class='src'>{linkify('جدول sadr (مبيعات)/سطر فاتورة 2332 · سطر 4154')}</td></tr>"
    t += f"<tr class='grand'><td>الإجمالي المنقول للمصروفات</td><td class='n'>{fmt(MOVED)}</td><td></td></tr></table>"
    t += '<div class="note">الكميات بالأطنان كما هي في المصدر (تشمل كميات الضرب التشغيلية)، التعديل على القيم المالية فقط.</div>'
    t += '<div class="note">أجور عمال التفريغ 88,579 ج بند بتحفظ (م9) خارج المجموع لاحتمال تداخله مع مصروفات البرنامج.</div>'
    # nonton sell
    t += '<h2>بنود البيع بلا وزن 656,943 ج (م4)</h2><table><tr><th>البند</th><th class="n">القيمة</th><th>المرجع</th></tr>'
    for name, d in MP['SELL_NT']:
        t += f"<tr><td>{name} ({d['rows']} سطر)</td><td class='n'>{fmt(d['value'])}</td><td class='src'>{linkify(d['ref'])}</td></tr>"
    t += '</table><h2>بنود الشراء بلا وزن 39,566,775 ج (م5)</h2><table><tr><th>البند</th><th class="n">القيمة</th><th>المرجع</th></tr>'
    for name, d in MP['BUY_NT'][:7]:
        t += f"<tr><td>{name} ({d['rows']} سطر)</td><td class='n'>{fmt(d['value'])}</td><td class='src'>{linkify(d['ref'])}</td></tr>"
    t += '</table><div class="note">جدول السنين بأرقام المصدر قبل إعادة التبويب.</div>'
    # years
    t += '<h2>السنين</h2><table><tr><th>السنة</th><th class="n">شراء (طن)</th><th class="n">بيع (طن)</th><th class="n">شراء (ج)</th><th class="n">بيع (ج)</th></tr>'
    for y, bt, st, bv, sv, _ in MP['YEARS']:
        t += f'<tr><td class="c">{y}</td><td class="n">{bt:,.1f}</td><td class="n">{st:,.1f}</td><td class="n">{fmt(bv)}</td><td class="n">{fmt(sv)}</td></tr>'
    t += '</table>'
    # program top items
    t += '<h2>أكبر بنود مصروفات البرنامج (م8)</h2><table><tr><th>البند</th><th class="n">سندات</th><th class="n">القيمة</th></tr>'
    for it in MP['PROG_ITEMS'][:10]:
        t += f"<tr><td>{it['item']}</td><td class='n'>{it['rows']}</td><td class='n'>{fmt(it['value'])}</td></tr>"
    t += '</table>'
    # mill fee treatment: internal to owners
    wb2 = openpyxl.load_workbook(NAT + 'أجرة المضرب والناولون - مضرب تعلب.xlsx', read_only=True, data_only=True)
    summ = {}
    for r in wb2['الملخص'].iter_rows(min_row=4, values_only=True):
        if r[1] and isinstance(r[2], (int, float)): summ[str(r[1]).strip()] = r[2]
    wb2.close()
    mill = summ.get('إجمالي أجرة المضرب', 22420107.44)
    naw = summ.get('إجمالي الناولون', 2861494.85)
    khaled = summ.get('إيجار مضرب خالد تعلب (14 سحبة)', 0)
    adj = nNET + mill + naw
    t += '<h2>معالجة أجرة المضرب والناولون (المضرب ملك الملاك)</h2>'
    t += '<div class="note">أجرة المضرب والناولون داخلة في الأرباح لأن المضرب ملك الملاك، وتكلفة المضرب خارجة من المصروفات لتسهيل الحساب. التفصيل الكامل في <a class="ref" href="2-اجرة-المضرب-والناولون.html">تقرير أجرة المضرب والناولون</a>.</div>'
    t += '<table><tr><th>البند</th><th class="n">القيمة (ج)</th></tr>'
    t += f"<tr><td>صافي الربح (بعد كل المصروفات)</td><td class='n'>{fmt(nNET)}</td></tr>"
    t += f"<tr><td>يرد: أجرة المضرب (داخلية)</td><td class='n'>{fmt(mill)}</td></tr>"
    t += f"<tr><td>يرد: الناولون (داخلي)</td><td class='n'>{fmt(naw)}</td></tr>"
    t += f"<tr class='grand'><td>الربح شامل الأجرة الداخلية</td><td class='n'>{fmt(adj)}</td></tr></table>"
    t += f'<div class="note">إيجار مضرب خالد تعلب {fmt(khaled)} ج داخل ضمن أجرة المضرب المردودة أعلاه.</div>'
    # refs
    t += '<h2>المراجع (م1 إلى م18)</h2>'
    for code, title, srcs, tbl, loc, how in MONEY_REFS:
        t += f'<div class="refbox"><b class="code">{code}</b> <b>{title}</b><br><span class="src">المصدر: {srcs} · {tbl}</span><br><span class="how">{linkify(loc,srcs,tbl)}</span><br><span>{linkify(how,srcs,tbl)}</span></div>'
    return page('المصروفات مجمعة والأرباح - مضرب تعلب',
                'الفترة الأولى (ديسمبر 2023 إلى يونيو 2024) + الفترة الثانية (يناير 2024 إلى أكتوبر 2025) + البرنامج (حتى 22/8/2026)',
                cards + t)

# ================= R2 =================
def build_r2():
    wb = openpyxl.load_workbook(NAT + 'أجرة المضرب والناولون - مضرب تعلب.xlsx', read_only=True, data_only=True)
    ws = wb['الملخص']
    t = '<h2>الملخص</h2><table><tr><th>البند</th><th class="n">القيمة (ج)</th><th class="n">سطور</th><th>المصدر</th></tr>'
    for r in list(ws.iter_rows(min_row=4, values_only=True)):
        if not r[1]: continue
        name = str(r[1]); val = r[2]; n = r[3]; src = str(r[4] or '')
        if 'إجمالي' in name or 'المجموع' in name: cls = 'total'
        else: cls = ''
        t += f"<tr class='{cls}'><td>{name}</td><td class='n'>{val if val is None else fmt(val)}</td><td class='n'>{n or ''}</td><td class='src'>{linkify(src,'الحركة 2الحركة حركه تعديل الخزينة','')}</td></tr>"
    t += '</table>'
    for sh in ['أجرة المضرب تفصيلي', 'الناولون تفصيلي']:
        ws = wb[sh]
        t += f'<h2>{sh}</h2><table><tr><th>م</th><th>التاريخ</th><th>البيان</th><th class="n">الكمية</th><th class="n">السعر</th><th class="n">القيمة</th><th>المصدر</th></tr>'
        for r in ws.iter_rows(min_row=3, values_only=True):
            if r[0] is None: continue
            src = str(r[6] or '')
            t += f"<tr><td class='c'>{r[0]}</td><td>{r[1] or ''}</td><td>{r[2] or ''}</td><td class='n'>{r[3] or ''}</td><td class='n'>{r[4] or ''}</td><td class='n'>{fmt(r[5]) if r[5] else ''}</td><td class='src'>{linkify(src,'الحركة 2الحركة حركه تعديل الخزينة','')}</td></tr>"
        t += '</table>'
    wb.close()
    return page('أجرة المضرب والناولون - مضرب تعلب', 'تفصيلي كل سطر بمصدره، دوس على المرجع يفتح السطر الأصلي', t)

# ================= R3 =================
def build_r3():
    wb = openpyxl.load_workbook(NAT + 'حسابا الأطنان - مضرب تعلب.xlsx', read_only=True, data_only=True)
    t = ''
    for sh in ['الحساب الأول · الطن الكامل', 'الحساب الثاني · الشعير وعملياته']:
        ws = wb[sh]
        t += f'<h2>{sh}</h2><table>'
        for i, r in enumerate(ws.iter_rows(values_only=True)):
            cells = [(str(c) if c is not None else '') for c in r]
            if all(c == '' for c in cells): continue
            if i <= 3:
                t += '<tr>' + ''.join(f'<th>{linkify(c)}</th>' for c in cells if c != '') + '</tr>'
            else:
                row = cells
                if len(row) > 4: row[-1] = linkify(row[-1])
                cls = 'total' if 'إجمالي' in cells[0] or 'الفرق' in cells[0] else ''
                t += f"<tr class='{cls}'>" + ''.join(f'<td>{c}</td>' for c in row) + '</tr>'
        t += '</table>'
    ws = wb['المراجع']
    t += '<h2>المراجع</h2>'
    for r in list(ws.iter_rows(min_row=3, values_only=True)):
        if not r[0]: continue
        det = ' · '.join(str(c) for c in [r[1], r[2], r[3], r[4]] if c)
        t += f"<div class='refbox'><b class='code'>{r[0]}</b> {linkify(det, str(r[2]), str(r[3]))}</div>"
    wb.close()
    return page('حسابا الأطنان - مضرب تعلب', 'شراء 59,109.9 طن · بيع 57,659.2 طن · الفرق 1,450.7 طن', t)

# ================= R4 =================
def build_r4():
    wb = openpyxl.load_workbook(NAT + 'تقرير الأصناف للمحاسب - تعلب.xlsx', read_only=True, data_only=True)
    t = ''
    for sh, head_rows in [('الملخص في صفحة', 2), ('من أين جاء كل رقم', 2), ('السنين', 1), ('أهم الأصناف', 2), ('الملاحظات المفتوحة', 2)]:
        ws = wb[sh]
        rows = list(ws.iter_rows(values_only=True))
        t += f'<h2>{sh}</h2><table>'
        for i, r in enumerate(rows):
            cells = [(str(c)[:90] if c is not None else '') for c in r]
            while cells and cells[-1] == '': cells.pop()
            if not cells: continue
            if i < head_rows:
                t += '<tr>' + ''.join(f'<th>{c}</th>' for c in cells) + '</tr>'
            else:
                cells = [linkify(numfmt(c)) for c in cells]
                cls = 'total' if any('الإجمالي' in c or 'النهائي' in c for c in cells) else ''
                t += f"<tr class='{cls}'>" + ''.join(f'<td>{c}</td>' for c in cells) + '</tr>'
        t += '</table>'
    t += '<h2>المراجع (1 إلى 18)</h2>'
    for title, srcs, tbl, loc, how in ACCT_REFS:
        t += f"<div class='refbox'><b>{title}</b><br><span class='src'>المصدر: {srcs} · {tbl}</span><br><span class='how'>{linkify(loc,srcs,tbl)}</span><br><span>{linkify(how,srcs,tbl)}</span></div>"
    wb.close()
    return page('تقرير الأصناف للمحاسب - تعلب', 'شراء 59,109.9 طن · بيع 57,659.2 طن · الفرق 1,450.7 طن', t)

builders = {'1-المصروفات-والارباح.html': build_r1, '2-اجرة-المضرب-والناولون.html': build_r2,
            '3-حسابا-الاطنان.html': build_r3, '4-تقرير-الاصناف-للمحاسب.html': build_r4}
for fn, b in builders.items():
    h = b().replace('<table>', '<div class="tbl"><table>').replace('</table>', '</table></div>')
    open(OUT + fn, 'w', encoding='utf-8').write(h)
    print('wrote', fn)
