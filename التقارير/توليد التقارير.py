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
    FULL = 25281602.29; KHALED = 2751908.94; EXN = 74786.25  # البند = الضرب كامل (بقرار المالك)
    nS = sell['total'] - 10000.0 - 1186750.0 + 283800.0; nB = buy['total'] - 10535500.0 - 2742708.6 - KHALED - 7851821.0  # +مرتجع مكرر مؤكد (سند 47/48 سطرين + 70/71، إدخال متتالي بثواني) يرد للبيع؛ -6 سطور شراء ف1 (3069-3075) إجماليها 10× الصحيح
    nE = MP['TOTAL_EXP'] + FULL; nNET = (nS - nB) - nE
    exp_e1 = MF.get('مصروفات_اكسل_ف1', MF.get('exp_e1', {}))
    cards = f"""<div class="cards">
<div class="card"><div class="v">{fmt(nS)}</div><div class="l">إجمالي البيع (ج)</div></div>
<div class="card"><div class="v">{fmt(nB)}</div><div class="l">إجمالي الشراء بعد نقل الأجرة (ج)</div></div>
<div class="card"><div class="v">{fmt(nS-nB)}</div><div class="l">الربح قبل المصروفات (ج)</div></div>
<div class="card"><div class="v">{fmt(nE - FULL)}</div><div class="l">إجمالي المصروفات (بدون بند الأجرة)</div></div>
<div class="card"><div class="v">{fmt(FULL)}</div><div class="l">بند الضرب والناولون (لوحده)</div></div>
<div class="card"><div class="v">{fmt(nNET)}</div><div class="l">صافي الربح بعد النقل (ج)</div></div>
<div class="card"><div class="v">{abs(sell['tons']-buy['tons']):,.1f}</div><div class="l">مخزون آخر المدة (طن)</div></div>
</div>"""
    exp_rows = [
        ('مصروفات الإكسل (مرة واحدة)', 844403, 'م6', xl_url('P1','مصروفات',17)),
        ('مرتبات وأجور العمال حتى 30/9/2025', 1949578, 'م7', xl_url('P2','مسير مرتبات')),
        ('مصروفات البرنامج 2/10/2025 إلى 22/8/2026', 4188879, 'م8', None),
    ]
    t = '<h2>المصروفات مجمعة</h2><table><tr><th>البند</th><th class="n">القيمة (ج)</th><th class="c">المرجع</th></tr>'
    for name, val, code, u in exp_rows:
        t += f'<tr><td>{name}</td><td class="n">{fmt(val)}</td><td class="c">{link(u, code)}</td></tr>'
    t += f'<tr class="grand"><td>إجمالي المصروفات (بدون بند الأجرة)</td><td class="n">{fmt(nE - FULL)}</td><td></td></tr></table>'
    t += '<h2>بند الأجرة الكاملة (تتخصم من الإجمالي وتتضاف على الأرباح)</h2><div class="note">الأجرة الكاملة بند لوحدها: خارج الشراء بلا وزن وخارج البيع، تتخصم من الإجمالي كمصروف ثم ترد كربح لأن المضرب ملك الملاك.</div><table><tr><th>المكون</th><th class="n">القيمة (ج)</th><th>كان فين</th></tr>'
    t += f"<tr><td>ضرب أرز برنامج (6 سندات ward)</td><td class='n'>{fmt(10535500.0)}</td><td class='src'>{linkify('جدول ward (مشتريات)/سطر فاتورة 1864 · سطر 2054')}</td></tr>"
    t += f"<tr><td>نولون شراء برنامج (947 سند ward)</td><td class='n'>{fmt(2742708.6)}</td><td class='src'>{linkify('جدول ward (مشتريات)/سطر فاتورة 1008 · سطر 1115')}</td></tr>"
    t += f"<tr><td>إيجار مضرب خالد تعلب (14 سحبة إكسل ف1)</td><td class='n'>{fmt(KHALED)}</td><td>كان في الشراء بلا وزن (اتشال)</td></tr>"
    t += f"<tr><td>نولون بيع برنامج (7 سندات sadr) - في البيع كإيراد وداخل في البند</td><td class='n'>{fmt(44000.0)}</td><td class='src'>{linkify('جدول sadr (مبيعات)/سطر فاتورة 2332 · سطر 4154')}</td></tr>"
    t += f"<tr><td>نولون ملفات الإكسل (13 سطر)</td><td class='n'>{fmt(EXN)}</td><td>في البيع كإيراد وداخل في البند</td></tr>"
    t += f"<tr><td>حساب المضرب سحبة 6:29 ومقطوعات (20)</td><td class='n'>{fmt(8445400.0)}</td><td>داخل في البند والرد (التزام دفتري لتشغيل المضرب)</td></tr>"
    t += f"<tr><td>أجرة مضرب 650 للطن (2)</td><td class='n'>{fmt(683598.5)}</td><td>خارج الإجماليات (داخل في البند والرد)</td></tr>"
    t += f"<tr><td>أجرة مضرب في مصروفات البرنامج (5)</td><td class='n'>{fmt(3700.0)}</td><td>داخل مصروفات البرنامج (م8) ومكررة في البند بفرق 3,700</td></tr>"
    t += f"<tr class='grand'><td>الأجرة الكاملة</td><td class='n'>{fmt(FULL)}</td><td></td></tr></table>"
    t += '<div class="note">الكميات بالأطنان كما هي في المصدر (تشمل كميات الضرب التشغيلية)، التعديل على القيم المالية فقط.</div>'

    t += '<div class="note">أجور عمال التفريغ 88,579 ج بند بتحفظ (م9) خارج المجموع لاحتمال تداخله مع مصروفات البرنامج.</div>'
    # nonton sell
    t += '<h2>بنود البيع بلا وزن 656,943 ج (م4)</h2><div class="note">ملحوظة: منها 10,000 ج (مبيعات سطر 2708) تحصيل خزنة وليست إيرادا، مخصومة من إجمالي البيع فوق.</div><table><tr><th>البند</th><th class="n">القيمة</th><th>المرجع</th></tr>'
    for name, d in MP['SELL_NT']:
        t += f"<tr><td>{name} ({d['rows']} سطر)</td><td class='n'>{fmt(d['value'])}</td><td class='src'>{linkify(d['ref'])}</td></tr>"
    t += '</table><div class="note">جدول المصدر قبل إعادة التبويب: بنود الأجرة (ضرب أرز 10,535,500 + نولون 2,742,709 + خالد 2,751,909) اتشالت من الشراء وبقت في بند الأجرة الكاملة وحده.</div>'
    NTB = sum(d['value'] for n2, d in MP['BUY_NT'] if norm_ar(n2) not in ('ضرب ارز', 'نولون'))
    t += f'<h2>بنود الشراء بلا وزن {fmt(NTB)} ج (م5، بعد شيل الأجرة)</h2><table><tr><th>البند</th><th class="n">القيمة</th><th>المرجع</th></tr>'
    for name, d in MP['BUY_NT']:
        if norm_ar(name) in ('ضرب ارز', 'نولون'): continue
        t += f"<tr><td>{name} ({d['rows']} سطر)</td><td class='n'>{fmt(d['value'])}</td><td class='src'>{linkify(d['ref'])}</td></tr>"
    t += '</table><div class="note">جدول السنين بأرقام المصدر قبل إعادة التبويب.</div>'
    # years
    t += '<h2>السنين</h2><table><tr><th>السنة</th><th class="n">شراء (طن)</th><th class="n">بيع (طن)</th><th class="n">شراء (ج)</th><th class="n">بيع (ج)</th></tr>'
    for y, bt, st, bv, sv, _ in MP['YEARS']:
        t += f'<tr><td class="c">{y}</td><td class="n">{bt:,.1f}</td><td class="n">{st:,.1f}</td><td class="n">{fmt(bv)}</td><td class="n">{fmt(sv)}</td></tr>'
    t += '</table>'
    t += '<h2>إجمالي الشكاير</h2><table><tr><th>البند</th><th class="n">سندات</th><th class="n">القيمة (ج)</th><th>المصدر</th></tr>'
    t += f"<tr><td>مشتريات إكسل فترة أولى (صنف شكاير بكميات)</td><td class='n'>43</td><td class='n'>{fmt(1324058)}</td><td>{link(DB + '?s=4', 'مشتريات وموردين ف1')}</td></tr>"
    t += f"<tr><td>مشتريات إكسل فترة ثانية (صنف شكاير بكميات)</td><td class='n'>166</td><td class='n'>{fmt(6349500)}</td><td>{link(DB + '?s=28', 'مشتريات وموردين ف2')}</td></tr>"
    t += f"<tr><td>مشتريات برنامج (ward ومرتجعات، 17 صنف)</td><td class='n'>127</td><td class='n'>{fmt(5656720)}</td><td>{link(DB + '?s=124', 'جدول ward')}</td></tr>"
    t += f"<tr><td>مصروفات إكسل لوازم إنتاج (نفس السطور في الملفين، محسوبة مرة واحدة)</td><td class='n'>9</td><td class='n'>{fmt(147656)}</td><td>{link(DB + '?s=3', 'مصروفات ف1')}</td></tr>"
    t += f"<tr class='grand'><td>إجمالي الشكاير شراء ومصروفات</td><td class='n'>345</td><td class='n'>{fmt(1324058 + 6349500 + 5656720 + 147656)}</td><td></td></tr></table>"
    t += f"<div class=\"note\">مبيعات شكاير: إكسل {fmt(285956)} ج ({link(DB + '?s=6', 'ف1')} + {link(DB + '?s=31', 'ف2')}) + برنامج {fmt(101625)} ج ({link(DB + '?s=123', 'جدول sadr')}).</div>"
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
    adj = nNET + FULL
    t += '<h2>معالجة أجرة المضرب والناولون (المضرب ملك الملاك)</h2>'
    t += '<div class="note">أجرة المضرب والناولون داخلة في الأرباح لأن المضرب ملك الملاك، وتكلفة المضرب خارجة من المصروفات لتسهيل الحساب. التفصيل الكامل في <a class="ref" href="2-اجرة-المضرب-والناولون.html">تقرير أجرة المضرب والناولون</a>.</div>'
    t += '<table><tr><th>البند</th><th class="n">القيمة (ج)</th></tr>'
    t += f"<tr><td>صافي الربح (بعد كل المصروفات)</td><td class='n'>{fmt(nNET)}</td></tr>"
    t += f"<tr><td>ترد: الأجرة الكاملة (بند واحد)</td><td class='n'>{fmt(FULL)}</td></tr>"
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

# ================= R5: purchase totals by item =================
def norm_ar(s):
    s = str(s or '').strip()
    s = s.replace('أ', 'ا').replace('إ', 'ا').replace('آ', 'ا').replace('ة', 'ه').replace('ى', 'ي')
    return ' '.join(s.split())

def build_r5():
    global R5GT
    TR = json.load(open('C:/Users/engmh/alrawda_check/tonrows.json', encoding='utf-8'))
    MD = json.load(open('C:/Users/engmh/alrawda_check/money_data.json', encoding='utf-8'))
    SKIP = {'ضرب ارز', 'نولون'}  # منقولة للمصروفات (م8-أجرة)، ليست شراء
    agg = {}
    def add(key, disp, src, rows, qty, val, unit=''):
        a = agg.setdefault(key, {'disp': disp, 'P1': [0, 0.0, 0.0, ''], 'P2': [0, 0.0, 0.0, ''], 'PG': [0, 0.0, 0.0, '']})
        c = a[src]; c[0] += rows; c[1] += qty; c[2] += val
        if unit: c[3] = unit
    for r in TR['ex_rows']:
        if r.get('sell'): continue
        it, grp = str(r.get('item') or ''), str(r.get('grp') or '')
        if not it: continue
        src = 'P2' if '2الحركة' in str(r.get('ref', '')) else 'P1'
        disp = it if not grp or grp in it or it in grp else f'{it} ({grp})'
        add('EX:' + norm_ar(it), disp, src, 1, r.get('qty') or 0, r.get('value') or 0, r.get('unit', ''))
    for r in TR['pg_rows']:
        if r.get('sell'): continue
        it = str(r.get('item') or '')
        if not it: continue
        grp = str(r.get('grp') or '')
        disp = it if not grp or grp in it or it in grp else f'{it} ({grp})'
        add('PG:' + norm_ar(it), disp, 'PG', 1, r.get('qty') or 0, r.get('value') or 0, r.get('unit', ''))
    MF = json.load(open('C:/Users/engmh/alrawda_check/money_final.json', encoding='utf-8'))
    UNSF = 'C:/Users/engmh/OneDrive/Documents/ghazali check/نتيجة/تفصيلي البيع والشراء بالمراجع - مضرب تعلب.xlsx'
    uwb = openpyxl.load_workbook(UNSF, read_only=True, data_only=True)
    uw = uwb['شراء كامل']
    for r in uw.iter_rows(min_row=5, values_only=True):
        if 'لم يُسجل' not in str(r[11] or ''): continue
        it = str(r[3] or '')
        if not it: continue
        v = r[9] if isinstance(r[9], (int, float)) else 0
        q = r[7] if isinstance(r[7], (int, float)) else 0
        if v > 0:
            add('EX:' + norm_ar(it), it, 'P2', 1, q, v, str(r[6] or ''))
    uwb.close()
    per = {}
    for f, src in [('C:/Users/engmh/OneDrive/Documents/ghazali check/new/الفترة الاولي/الحركة (بدون باسورد).xlsm', 'P1'),
                   ('C:/Users/engmh/OneDrive/Documents/ghazali check/new/الفترة الثانية/2الحركة (بدون باسورد).xlsm', 'P2')]:
        wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
        for r in wb['مشتريات وموردين'].iter_rows(min_row=3, values_only=True):
            if any(k in str(r[3] or '') for k in ['شكاير', 'شيكارة', 'شكارة']):
                v = r[6] if isinstance(r[6], (int, float)) else 0
                q = r[4] if isinstance(r[4], (int, float)) else 0
                if v > 0:
                    d = per.setdefault(src, [0, 0.0, 0.0])
                    d[0] += 1; d[1] += q; d[2] += v
        wb.close()
    for src, (n, q, v) in per.items():
        add('EXSH', 'شكاير (أصناف الإكسل بكميات - خارج سطور المراجعة)', src, n, q, v, 'شيكارة')
    def cell(c, sid):
        if not c[0]: return '<td></td>'
        q = f"{c[1]:,.1f} {c[3]}".strip() if c[1] else ''
        return f"<td class='n'>{q}<br><b>{fmt(c[2])}</b> ج<br>{link(DB + f'?s={sid}', 'المصدر')}</td>"
    t = '<table><tr><th>الصنف</th><th>فترة أولى (كمية / قيمة)</th><th>فترة ثانية (كمية / قيمة)</th><th>البرنامج (كمية / قيمة)</th><th class="n">إجمالي القيمة (ج)</th></tr>'
    gt = 0
    for k, a in sorted(agg.items(), key=lambda kv: -(kv[1]['P1'][2] + kv[1]['P2'][2] + kv[1]['PG'][2])):
        tot = a['P1'][2] + a['P2'][2] + a['PG'][2]
        gt += tot
        t += f"<tr><td>{a['disp']}</td>{cell(a['P1'], 4)}{cell(a['P2'], 28)}{cell(a['PG'], 124)}<td class='n'><b>{fmt(tot)}</b></td></tr>"
    ret_val = MF['buy']['src']['مرتجع (خصم)'] if isinstance(MF.get('buy',{}).get('src',{}).get('مرتجع (خصم)'), (int,float)) else -32810.9
    if ret_val:
        t += f"<tr><td>مرتجع شراء (خصم) — 21 سطر return_w</td><td></td><td></td><td class='n'><b>{fmt(ret_val)}</b> ج<br>{link(DB + '?s=124', 'جدول return_w')}</td><td class='n'><b>{fmt(ret_val)}</b></td></tr>"
        gt += ret_val
    t += f"<tr><td>تصحيح: 6 سطور شراء ف1 (3069-3075) إجماليها 10× الصحيح (22-06-2024)</td><td></td><td class='n'><b>{fmt(-7851821.0)}</b> ج</td><td></td><td class='n'><b>{fmt(-7851821.0)}</b></td></tr>"
    gt += -7851821.0
    R5GT = gt
    t += f"<tr class='grand'><td>الإجمالي (بعد خصم المرتجع والتصحيح)</td><td></td><td></td><td></td><td class='n'>{fmt(gt)}</td></tr></table>"
    ded = 10535500.0 + 2742708.6 + 2751908.94 + 7673558.0
    uni = gt - ded
    t += f"""<table><tr><th>بند الأجرة والشكاير (خارج الشراء — في بند الضرب والناولون لوحده)</th><th class="n">القيمة (ج)</th></tr>
<tr><td>ضرب أرز — 6 سطور ward فاتورة 1864 (أجرة وليست شراء)</td><td class='n'><b>{fmt(10535500.0)}</b></td></tr>
<tr><td>نولون — 947 نقلة ward فاتورة 1008 (أجرة وليست شراء)</td><td class='n'><b>{fmt(2742708.6)}</b></td></tr>
<tr><td>إيجار مضرب خالد — 14 سحبة إكسل (أجرة وليست شراء)</td><td class='n'><b>{fmt(2751908.94)}</b></td></tr>
<tr><td>شكاير الإكسل — 209 سندات (بند مستقل للعلم)</td><td class='n'><b>{fmt(7673558.0)}</b></td></tr>
<tr class='grand'><td>الصافي الموحد (مطابق إجمالي الشراء في صفحة المصروفات والأرباح)</td><td class='n'>{fmt(uni)}</td></tr></table>"""
    t += '<div class="note">من سطور المراجعة المعتمدة + شكاير الإكسل (209 سندات) التي لم تدخل سطورها. شاملة "إكسل لم يُسجل في البرنامج" (141 سنداً مدمجة في الفترة الثانية). المرتجع شراء −32,811 ج مخصوم (21 سطر). ضرب الأرز والنولون خارج الجدول لأنهما في بند الأجرة وليسا شراء. شكاير: إكسل فترة أولى+ثانية + برنامج ward.</div>'
    return page('إجمالي المشتريات بالأصناف - مضرب تعلب', 'كل صنف اشتريته: فترة أولى + فترة ثانية + برنامج', t)

# ================= R6: sales totals by item =================
def build_r6():
    TR = json.load(open('C:/Users/engmh/alrawda_check/tonrows.json', encoding='utf-8'))
    MF = json.load(open('C:/Users/engmh/alrawda_check/money_final.json', encoding='utf-8'))
    agg = {}
    def add(key, disp, src, rows, qty, val, unit=''):
        a = agg.setdefault(key, {'disp': disp, 'P1': [0, 0.0, 0.0, ''], 'P2': [0, 0.0, 0.0, ''], 'PG': [0, 0.0, 0.0, '']})
        c = a[src]; c[0] += rows; c[1] += qty; c[2] += val
        if unit: c[3] = unit
    for r in TR['ex_rows']:
        if not r.get('sell'): continue
        it, grp = str(r.get('item') or ''), str(r.get('grp') or '')
        if not it: continue
        src = 'P2' if '2الحركة' in str(r.get('ref', '')) else 'P1'
        disp = it if not grp or grp in it or it in grp else f'{it} ({grp})'
        add('EX:' + norm_ar(it), disp, src, 1, r.get('qty') or 0, r.get('value') or 0, r.get('unit', ''))
    for r in TR['pg_rows']:
        if not r.get('sell'): continue
        it = str(r.get('item') or '')
        if not it: continue
        grp = str(r.get('grp') or '')
        disp = it if not grp or grp in it or it in grp else f'{it} ({grp})'
        add('PG:' + norm_ar(it), disp, 'PG', 1, r.get('qty') or 0, r.get('value') or 0, r.get('unit', ''))
    UNSF = 'C:/Users/engmh/OneDrive/Documents/ghazali check/نتيجة/تفصيلي البيع والشراء بالمراجع - مضرب تعلب.xlsx'
    uwb = openpyxl.load_workbook(UNSF, read_only=True, data_only=True)
    uw = uwb['بيع كامل']
    for r in uw.iter_rows(min_row=5, values_only=True):
        if 'لم يُسجل' not in str(r[11] or ''): continue
        it = str(r[3] or '')
        if not it: continue
        v = r[9] if isinstance(r[9], (int, float)) else 0
        q = r[7] if isinstance(r[7], (int, float)) else 0
        if v > 0:
            add('EX:' + norm_ar(it), it, 'P2', 1, q, v, str(r[6] or ''))
    uwb.close()
    def cell(c, sid):
        if not c[0]: return '<td></td>'
        q = f"{c[1]:,.1f} {c[3]}".strip() if c[1] else ''
        return f"<td class='n'>{q}<br><b>{fmt(c[2])}</b> ج<br>{link(DB + f'?s={sid}', 'المصدر')}</td>"
    t = '<table><tr><th>الصنف</th><th>فترة أولى (كمية / قيمة)</th><th>فترة ثانية (كمية / قيمة)</th><th>البرنامج (كمية / قيمة)</th><th class="n">إجمالي القيمة (ج)</th></tr>'
    gt = 0
    for k, a in sorted(agg.items(), key=lambda kv: -(kv[1]['P1'][2] + kv[1]['P2'][2] + kv[1]['PG'][2])):
        tot = a['P1'][2] + a['P2'][2] + a['PG'][2]
        gt += tot
        t += f"<tr><td>{a['disp']}</td>{cell(a['P1'], 6)}{cell(a['P2'], 31)}{cell(a['PG'], 123)}<td class='n'><b>{fmt(tot)}</b></td></tr>"
    ret_val = MF['sell']['src']['مرتجع (خصم)'] if isinstance(MF.get('sell',{}).get('src',{}).get('مرتجع (خصم)'), (int,float)) else -14282830.7
    if ret_val:
        t += f"<tr><td>مرتجع بيع (خصم)</td><td></td><td></td><td class='n'><b>{fmt(ret_val)}</b> ج<br>{link(DB + '?s=123', 'جدول sadr')}</td><td class='n'><b>{fmt(ret_val)}</b></td></tr>"
        gt += ret_val
    t += f"<tr><td>تحصيل خزنة ليس إيرادا (مبيعات سطر 2708 - نولون عيد الصعيدي)</td><td></td><td class='n'><b>{fmt(-10000.0)}</b> ج</td><td></td><td class='n'><b>{fmt(-10000.0)}</b></td></tr>"
    gt += -10000.0
    t += f"<tr><td>تصحيح خطأ إدخال (مبيعات ف1 سطر 5302: 500كجم أرز عريض بـ1,200,000، الصحيح 13,250 بسعر نفس اليوم 26.5)</td><td class='n'><b>{fmt(-1186750.0)}</b> ج</td><td></td><td></td><td class='n'><b>{fmt(-1186750.0)}</b></td></tr>"
    gt += -1186750.0
    t += f"<tr><td>مرتجع مكرر يُرد للبيع (سند 47/48 سطرين + سند 70/71 — إدخال متتالي بثواني)</td><td></td><td></td><td class='n'><b>{fmt(283800.0)}</b> ج</td><td class='n'><b>{fmt(283800.0)}</b></td></tr>"
    gt += 283800.0
    t += f"<tr class='grand'><td>الإجمالي (بعد خصم المرتجع والتحصيل ورد المكرر)</td><td></td><td></td><td></td><td class='n'>{fmt(gt)}</td></tr></table>"
    t += '<div class="note">من سطور المراجعة المعتمدة. شاملة "إكسل لم يُسجل في البرنامج" بيع (361 سنداً مدمجة في الفترة الثانية). المرتجع −14,282,831 ج مخصوم. التحصيل −10,000 ج مخصوم لأنه خزنة وليس إيرادا.</div>'
    return page('إجمالي المبيعات بالأصناف - مضرب تعلب', 'كل صنف بعته: فترة أولى + فترة ثانية + برنامج', t)

# ================= R7: إجمالي المنصرف =================
def build_r7():
    MF = json.load(open('C:/Users/engmh/alrawda_check/money_final.json', encoding='utf-8'))
    gt = R5GT  # إجمالي المشتريات بعد المرتجع والتصحيح
    darb, naw, khaled, exsh = 10535500.0, 2742708.6, 2751908.94, 7673558.0
    A = gt - darb - naw - exsh - khaled  # مشتريات البضاعة (بعد فصل خالد لبند الضرب)
    oil, sams = 19709961.24, 921885.0
    grain = A - oil - sams  # حبوب وخامات (صافي شامل كل التسويات)
    FULL = 25281602.29
    e1, p1, p2, pg = 844403.0, 105065.0, 1844512.5, 4188879.0
    exp = e1 + p1 + p2 + pg
    grand = A + exsh + exp  # البند خارج المنصرف — يرد على الأرباح
    r = lambda n, v: f"<tr><td>{n}</td><td class='n'><b>{fmt(v)}</b></td></tr>"
    t = '<table><tr><th>بند المنصرف</th><th class="n">القيمة (ج)</th></tr>'
    t += r('حبوب وخامات طنّية (صافي شامل كل التسويات)', grain)
    t += r('زيت (مستبعد من الربح بقرار الملاك)', oil)
    t += r('سمسرة ولوازم شراء (845,180 + 59,000 + 9,000 + 6,000 + 1,425 + 1,280)', sams)
    t += f"<tr class='grand'><td>مشتريات البضاعة</td><td class='n'>{fmt(A)}</td></tr>"
    t += r('شكاير الإكسل — 209 سندات (خارج سطور المراجعة)', exsh)
    t += f"<tr><td>إجمالي الشكاير للعلم (برنامج 5,656,720 + إكسل 7,673,558 + مصروفات 147,656)</td><td class='n'><b>{fmt(13477934.0)}</b> — معلومة لا تضاف</td></tr>"
    t += f"<tr><td>بند الضرب والناولون (خارج المنصرف — يرد على الأرباح): ضرب 10,535,500 + نولون 2,742,709 + خالد 2,751,909 + حساب مضرب 8,445,400 + 650/طن 683,599 + نولون إكسل 74,786 + نولون بيع 44,000 + فرق 3,700 = <b>{fmt(FULL)}</b></td><td class='n'>معلومة لا تضاف</td></tr>"
    t += r('مصروفات إكسل', e1) + r('مرتبات فترة أولى', p1) + r('مرتبات فترة ثانية', p2) + r('مصروفات البرنامج', pg)
    t += f"<tr class='grand'><td>مصروفات التشغيل</td><td class='n'>{fmt(exp)}</td></tr>"
    t += f"<tr class='grand'><td>المجموع الكلي للمنصرف</td><td class='n'>{fmt(grand)}</td></tr></table>"
    t += '<div class="note">كل جنيه اتصرف: بضاعة + شكاير إكسل + أجرة + مصروفات. الحبوب محسوبة كصافي (إجمالي المشتريات − الزيت − السمسرة − الأجرة − شكاير الإكسل) فالصفحة موزونة بالجنيه.</div>'
    return page('إجمالي المنصرف - مضرب تعلب', 'كل حاجة اتصرف فيها فلوس: خامات + شكاير + أجرة + مصروفات', t)

builders = {'1-المصروفات-والارباح.html': build_r1, '2-اجرة-المضرب-والناولون.html': build_r2,\
            '3-حسابا-الاطنان.html': build_r3, '4-تقرير-الاصناف-للمحاسب.html': build_r4,\
            '5-إجمالي-المشتريات-بالأصناف.html': build_r5, '6-إجمالي-المبيعات-بالأصناف.html': build_r6,\
            '7-إجمالي-المنصرف.html': build_r7}
for fn, b in builders.items():
    h = b().replace('<table>', '<div class="tbl"><table>').replace('</table>', '</table></div>')
    open(OUT + fn, 'w', encoding='utf-8').write(h)
    print('wrote', fn)
