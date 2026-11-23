# -*- coding: utf-8 -*-
"""توليد صفحات الفترة - 2 من بيانات قاعدة البيانات الموجودة في المجلد data."""
import os, json, io, math

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, 'data')
OUT = os.path.join(ROOT, 'period-2')
os.makedirs(OUT, exist_ok=True)


def load(sid):
    with io.open(os.path.join(DATA, f'{sid}.json'), encoding='utf-8') as f:
        return json.load(f)


def num(v):
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        try:
            return float(v.replace(',', '').strip())
        except ValueError:
            return 0.0
    return 0.0


def s(v):
    return '' if v is None else str(v).strip()


def money(v, dp=2):
    return f'{v:,.{dp}f}'


CSS = """
:root{--bg:#050814;--bg2:#070d1e;--tx:#eef1f6;--mut:rgba(238,241,246,.68);--card:#111a2e;--bd:rgba(255,255,255,.09);--gold:#E8C547;--gold-dim:rgba(232,197,71,.14)}
*{box-sizing:border-box}
body{margin:0;background:linear-gradient(180deg,var(--bg),var(--bg2));color:var(--tx);font-family:'Segoe UI',Tahoma,Arial,sans-serif;direction:rtl;font-size:14px;line-height:1.6}
a{color:var(--gold);text-decoration:none}
a:hover{text-decoration:underline}
.wrap{max-width:1300px;margin:0 auto;padding:20px 16px 60px}
header{background:linear-gradient(135deg,#1a2540,#0f1a33);border:1px solid rgba(232,197,71,.28);border-radius:14px;padding:22px 24px;margin-bottom:18px;box-shadow:0 8px 26px rgba(0,0,0,.35)}
h1{margin:0 0 6px;font-size:26px;color:var(--gold)}
h2{font-size:19px;margin:30px 0 10px;color:var(--gold);border-bottom:1px solid var(--bd);padding-bottom:7px}
.sub{color:var(--mut);font-size:13px}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:12px;margin:16px 0}
.card{background:var(--card);border:1px solid var(--bd);border-radius:12px;padding:14px}
.card .l{color:var(--mut);font-size:12px}
.card .v{font-size:21px;font-weight:700;color:var(--gold);direction:ltr;unicode-bidi:isolate;display:block}
.links{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:14px}
.linkcard{background:var(--card);border:1px solid var(--bd);border-radius:12px;padding:20px;display:block;color:var(--tx);transition:.15s}
.linkcard:hover{border-color:var(--gold);text-decoration:none;transform:translateY(-2px)}
.linkcard b{display:block;color:var(--gold);font-size:19px;margin-bottom:6px}
.bar{display:flex;gap:10px;flex-wrap:wrap;margin:12px 0;align-items:center}
input{background:#0b1222;color:var(--tx);border:1px solid var(--bd);border-radius:8px;padding:9px 12px;font-size:14px;min-width:260px}
button{background:var(--gold-dim);color:var(--gold);border:1px solid rgba(232,197,71,.35);border-radius:8px;padding:8px 14px;cursor:pointer;font-size:13px}
button:hover{background:rgba(232,197,71,.24)}
.tbox{background:var(--card);border:1px solid var(--bd);border-radius:12px;overflow:auto}
table{border-collapse:collapse;width:100%;font-size:12.5px}
th{background:rgba(255,255,255,.05);color:var(--gold);position:sticky;top:0;padding:9px 8px;text-align:right;white-space:nowrap;border-bottom:1px solid var(--bd);z-index:1}
td{padding:7px 8px;border-bottom:1px solid rgba(255,255,255,.05);vertical-align:top}
tr:nth-child(even) td{background:rgba(255,255,255,.015)}
tr:hover td{background:rgba(232,197,71,.07)}
td.n,th.n{text-align:left;direction:ltr;unicode-bidi:isolate;white-space:nowrap;font-variant-numeric:tabular-nums}
tr.sum td,tr.tot td{background:rgba(232,197,71,.12);font-weight:700}
.sumhdr{font-size:20px;color:var(--gold);font-weight:700;margin:30px 0 4px;padding-bottom:8px;border-bottom:1px solid rgba(232,197,71,.35)}
.ref{font-size:11px;background:var(--gold-dim);border:1px solid rgba(232,197,71,.25);border-radius:6px;padding:2px 7px;white-space:normal;display:inline-block}
.pg{display:flex;gap:8px;align-items:center;margin-top:10px;color:var(--mut);font-size:13px}
.note{background:rgba(232,197,71,.06);border:1px solid rgba(232,197,71,.2);border-right:3px solid var(--gold);border-radius:8px;padding:12px 16px;margin:16px 0;color:var(--mut);font-size:13px}
.back{display:inline-block;margin-bottom:14px;font-size:13px}
td.neg{color:#ff9b9b}
@media(max-width:600px){h1{font-size:21px}input{min-width:100%}th,td{padding:6px;font-size:11px}}
"""

JS = """
function mkTable(id, rows, cols, per){
  const box = document.getElementById(id+'-box');
  const tb = document.getElementById(id+'-tb');
  const q = document.getElementById(id+'-q');
  const info = document.getElementById(id+'-info');
  const prev = document.getElementById(id+'-prev');
  const next = document.getElementById(id+'-next');
  let page = 0, filt = rows;
  function fmt(v){
    if(v===null||v===undefined) return '';
    if(typeof v==='number') return v.toLocaleString('en-US',{maximumFractionDigits:2});
    return String(v).replace(/[<>]/g,'');
  }
  function render(){
    const start = page*per;
    const part = filt.slice(start,start+per);
    let html = '<thead><tr>'+cols.map(c=>'<th'+(c.n?' class="n"':'')+'>'+c.t+'</th>').join('')+'</tr></thead><tbody>';
    for(const r of part){
      html += '<tr>';
      cols.forEach(c=>{
        let v = r[c.k];
        let cls = c.n?' class="n"':'';
        if(typeof v==='number'&&v<0) cls+=' neg';
        if(c.ref){ html += '<td><a class="ref" target="_blank" href="../index.html?s='+r.sid+'&r='+r.ri+'">فتح المصدر</a></td>'; }
        else html += '<td'+cls+'>'+fmt(v)+'</td>';
      });
      html += '</tr>';
    }
    tb.innerHTML = html;
    info.textContent = (filt.length? start+1:0)+' - '+Math.min(start+per,filt.length)+' من '+filt.length.toLocaleString('en-US');
    prev.disabled = page===0;
    next.disabled = (page+1)*per>=filt.length;
  }
  q.addEventListener('input',()=>{
    const v=q.value.trim().toLowerCase();
    filt = v? rows.filter(r=>cols.some(c=>String(r[c.k]||'').toLowerCase().includes(v))) : rows;
    page=0; render();
  });
  prev.onclick=()=>{if(page>0){page--;render();}};
  next.onclick=()=>{if((page+1)*per<filt.length){page++;render();}};
  render();
}
"""


def page(title, subtitle, blocks):
    """blocks: list of (id, heading, note, columns, rows, per)"""
    h = [f'<!DOCTYPE html><html lang="ar" dir="rtl"><head><meta charset="utf-8">',
         '<meta name="viewport" content="width=device-width,initial-scale=1">',
         f'<title>{title} · الفترة - 2</title><style>{CSS}</style></head><body><div class="wrap">',
         '<a class="back" href="index.html">◀ رجوع للفترة - 2</a>',
         f'<header><h1>{title}</h1><div class="sub">{subtitle}</div></header>']
    for ident, heading, note, cols, rows, per in blocks:
        h.append(f'<h2>{heading}</h2>')
        if note:
            h.append(f'<div class="note">{note}</div>')
        h.append(f'<div class="bar"><input id="{ident}-q" placeholder="ابحث في الجدول..."><button id="{ident}-prev">▶ السابق</button><button id="{ident}-next">الالتالي ◀</button><span id="{ident}-info"></span></div>')
        h.append(f'<div class="tbox"><table id="{ident}-tb"></table></div>')
    h.append(f'<script>const D={json.dumps({b[0]: b[4] for b in blocks}, ensure_ascii=False, separators=(",",":"))};')
    h.append(f'const C={json.dumps({b[0]: b[3] for b in blocks}, ensure_ascii=False, separators=(",",":"))};')
    h.append('</script><script>' + JS)
    h.append('for(const k in D) mkTable(k,D[k],C[k],100);')
    h.append('</script></div></body></html>')
    return '\n'.join(h)


def write(name, text):
    with io.open(os.path.join(OUT, name), 'w', encoding='utf-8', newline='\n') as f:
        f.write(text)
    print(name, f'{len(text)/1024:.0f} KB')


# ============ شراء ============
d28 = load(28)
buy = []
for i, r in enumerate(d28['rows']):
    w, v = num(r[4]), num(r[6])
    txt = ' '.join(s(x) for x in r[:4])
    if w == 0 and v == 0 and 'مرتجع' not in txt:
        continue
    buy.append(dict(sid=28, ri=i, date=s(r[10])[:10], name=s(r[1]), byan=s(r[2]), item=s(r[3]),
                    kg=round(w, 3), price=round(num(r[5]), 2), val=round(v, 2),
                    draw=s(r[11])))
buy_qty = sum(x['kg'] for x in buy) / 1000
buy_val = sum(x['val'] for x in buy)
buy_ret = [x for x in buy if 'مرتجع' in (x['byan'] + x['name'])]
from collections import defaultdict
by_sup = defaultdict(float)
for x in buy:
    by_sup[x['name']] += x['val']
sup_rows = [dict(sid=28, ri=0, name=k, n=0, val=round(v, 2)) for k, v in sorted(by_sup.items(), key=lambda kv: -kv[1])[:25]]

BC = [dict(t='م', k='n'), dict(t='التاريخ', k='date'), dict(t='المورد', k='name'),
      dict(t='البيان', k='byan'), dict(t='الصنف', k='item'), dict(t='وزن كجم', k='kg', n=1),
      dict(t='السعر', k='price', n=1), dict(t='القيمة جنيه', k='val', n=1),
      dict(t='السحبة', k='draw'), dict(t='المصدر', k='ref', ref=1)]
SC = [dict(t='م', k='n'), dict(t='المورد', k='name'), dict(t='القيمة جنيه', k='val', n=1), dict(t='', k='ref', ref=1)]
for i, x in enumerate(buy, 1): x['n'] = i
for i, x in enumerate(sup_rows, 1): x['n'] = i

write('purchases.html', page(
    'شراء الفترة الثانية',
    'المصدر: 2الحركة (بدون باسورد).xlsm · شيت مشتريات وموردين · كل سطر مرتبط بالصف الأصلي في قاعدة البيانات',
    [('buy', f'تفاصيل المشتريات ({len(buy):,} سطر)',
      f'الوزن الإجمالي {money(buy_qty, 3)} طن · القيمة {money(buy_val)} جنيه · مرتجعات {len(buy_ret)} سطر. كل الأرقام من شيت مشتريات وموردين في ملف 2الحركة للفترة الثانية فقط، والمرتجعات ظاهرة ضمن السطور.',
      BC, buy, 100),
     ('sup', 'أعلى الموردين قيمة', 'مرتبة حسب إجمالي القيمة في نفس الشيت.', SC, sup_rows, 50)]))

# ============ بيع ============
d31 = load(31)
sell = []
for i, r in enumerate(d31['rows']):
    w, v = num(r[5]), num(r[7])
    txt = ' '.join(s(x) for x in r[1:5])
    if w == 0 and v == 0 and 'مرتجع' not in txt:
        continue
    sell.append(dict(sid=31, ri=i, date=s(r[11])[:10], name=s(r[1]), byan=s(r[2]), item=s(r[3]),
                     inv=s(r[4]), kg=round(w, 3), price=round(num(r[6]), 2), val=round(v, 2),
                     paid=round(num(r[8]), 2), rest=round(num(r[9]), 2), draw=s(r[12])))
sell_qty = sum(x['kg'] for x in sell) / 1000
sell_val = sum(x['val'] for x in sell)
sell_ret = [x for x in sell if 'مرتجع' in (x['byan'] + x['name'])]
by_cus = defaultdict(float)
for x in sell:
    by_cus[x['name']] += x['val']
cus_rows = [dict(sid=31, ri=0, name=k, n=0, val=round(v, 2)) for k, v in sorted(by_cus.items(), key=lambda kv: -kv[1])[:25]]
VC = [dict(t='م', k='n'), dict(t='التاريخ', k='date'), dict(t='العميل', k='name'),
      dict(t='البيان', k='byan'), dict(t='الصنف', k='item'), dict(t='وزن كجم', k='kg', n=1),
      dict(t='السعر', k='price', n=1), dict(t='القيمة جنيه', k='val', n=1),
      dict(t='السحبة', k='draw'), dict(t='المصدر', k='ref', ref=1)]
for i, x in enumerate(sell, 1): x['n'] = i
for i, x in enumerate(cus_rows, 1): x['n'] = i
write('sales.html', page(
    'بيع الفترة الثانية',
    'المصدر: 2الحركة (بدون باسورد).xlsm · شيت مبيعات وعملاء · كل سطر مرتبط بالصف الأصلي في قاعدة البيانات',
    [('sell', f'تفاصيل المبيعات ({len(sell):,} سطر)',
      f'الوزن الإجمالي {money(sell_qty, 3)} طن · القيمة {money(sell_val)} جنيه · مرتجعات {len(sell_ret)} سطر. كل الأرقام من شيت مبيعات وعملاء في ملف 2الحركة للفترة الثانية فقط، والمرتجعات ظاهرة ضمن السطور.',
      VC, sell, 100),
     ('cus', 'أعلى العملاء قيمة', 'مرتبة حسب إجمالي القيمة في نفس الشيت.', SC, cus_rows, 50)]))

# ============ ضرب وناولون ============
d34 = load(34)
mill = []
for i, r in enumerate(d34['rows']):
    txt = s(r[1])
    if not txt:
        continue
    if not any(k in txt for k in ('مضرب', 'اجرة', 'أجرة', 'سحبة', 'سحبه', 'ضرب', 'نولون', 'ناولون')):
        continue
    mill.append(dict(sid=34, ri=i, date=s(r[0])[:10], byan=txt,
                     debit=round(num(r[2]), 2), credit=round(num(r[3]), 2), bal=round(num(r[4]), 2)))
mill_amt = sum(x['credit'] for x in mill)

# بنود الضرب والناولون داخل مشتريات ومبيعات ف2 + الخزينة + البرنامج
freight = []
for sid, d, cols in [(28, d28, (1, 2, 3, 4, 6, 10, 11)),
                     (31, load(31), (1, 2, 3, 5, 7, 11, 12))]:
    for i, r in enumerate(d['rows']):
        txt = ' '.join(s(r[j]) for j in (1, 2, 3))
        if not any(k in txt for k in ('نولون', 'ناولون', 'ضرب أرز', 'ضرب الارز', 'اجرة مضرب')):
            continue
        freight.append(dict(sid=sid, ri=i, date=s(r[cols[5]])[:10], who=s(r[cols[0]]),
                            byan=s(r[cols[1]]), item=s(r[cols[2]]),
                            val=round(num(r[cols[4]]), 2), src='إكسل ف2'))
d26 = load(26)
for i, r in enumerate(d26['rows']):
    txt = s(r[2])
    if not any(k in txt for k in ('نولون', 'ناولون', 'اجرة مضرب', 'أجرة مضرب', 'ضرب أرز')):
        continue
    freight.append(dict(sid=26, ri=i, date=s(r[4])[:10], who='الخزينة', byan=txt,
                        item='', val=round(num(r[0]) + num(r[1]), 2), src='الخزينة ف2'))
d117 = load(117)
for i, r in enumerate(d117['rows']):
    if 'مضرب' not in ' '.join(s(x) for x in r):
        continue
    freight.append(dict(sid=117, ri=i, date=s(r[8])[:10], who='الخزينة Sheet1',
                        byan='أجرة مضرب 650 للطن', item='', val=round(num(r[2]), 2),
                        src='الخزينة Sheet1 ف2'))
freight.sort(key=lambda x: x['date'])
for i, x in enumerate(freight, 1): x['n'] = i

MC = [dict(t='م', k='n'), dict(t='التاريخ', k='date'), dict(t='البيان', k='byan'),
      dict(t='مدين', k='debit', n=1), dict(t='دائن', k='credit', n=1),
      dict(t='الرصيد', k='bal', n=1), dict(t='المصدر', k='ref', ref=1)]
FC = [dict(t='م', k='n'), dict(t='التاريخ', k='date'), dict(t='الجهة', k='who'),
      dict(t='البيان', k='byan'), dict(t='الصنف', k='item'),
      dict(t='القيمة جنيه', k='val', n=1), dict(t='الملف', k='src'), dict(t='المصدر', k='ref', ref=1)]
write('milling-freight.html', page(
    'ضرب وناولون الفترة الثانية',
    'المصادر: حركه تعديل · شيت حساب المضرب · 2الحركة ف2 · الخزينة ف2 · الخزينة Sheet1 · الفترة الثانية فقط',
    [('mill', f'حساب المضرب ({len(mill):,} سطر)',
      f'إجمالي الدائن {money(mill_amt)} جنيه، منه أجرة المضرب 8,445,400 جنيه. كل سطر يفتح الصف الأصلي في الفترة الثانية.',
      MC, mill, 100),
     ('fr', f'بنود الضرب والناولون ({len(freight):,} سطر)',
      f'بنود الفترة الثانية فقط: مشتريات ومبيعات ف2، الخزينة ف2، والخزينة Sheet1. تحفظ: سطر 30/9 في الخزينة 378,950 وسطر 30/9 في حساب المضرب 309,000 نفس التاريخ بمبلغين مختلفين.',
      FC, freight, 100)]))

# ============ مصروفات ============
d27 = load(27)
expsum = []
for i, r in enumerate(d27['rows']):
    if num(r[7]) == 0 or not s(r[6]):
        continue
    expsum.append(dict(sid=27, ri=i, n=0, name=s(r[6]), val=round(num(r[7]), 2)))
expsum_total = sum(x['val'] for x in expsum)
exp = []
for i, r in enumerate(d27['rows']):
    if num(r[0]) == 0 or not s(r[3]):
        continue
    exp.append(dict(sid=27, ri=i, date=s(r[3])[:10], typ=s(r[1]), byan=s(r[2]),
                    val=round(num(r[0]), 2)))
exp_total = sum(x['val'] for x in exp)
d44 = load(44)
pay = []
for i, r in enumerate(d44['rows']):
    v = num(r[7])
    if v == 0 or not s(r[8]):
        continue
    pay.append(dict(sid=44, ri=i, date=s(r[8])[:10], name=s(r[1]),
                    worker=s(r[2]), barley=s(r[3]), white=s(r[4]),
                    val=round(v, 2)))
pay_total = sum(x['val'] for x in pay)
for k, rows in ((0, expsum), (1, exp), (2, pay)):
    for i, x in enumerate(rows, 1): x['n'] = i
EC = [dict(t='م', k='n'), dict(t='البند', k='name'), dict(t='القيمة جنيه', k='val', n=1), dict(t='المصدر', k='ref', ref=1)]
XD = [dict(t='م', k='n'), dict(t='التاريخ', k='date'), dict(t='النوع', k='typ'),
      dict(t='البيان', k='byan'), dict(t='القيمة جنيه', k='val', n=1), dict(t='المصدر', k='ref', ref=1)]
PC = [dict(t='م', k='n'), dict(t='التاريخ', k='date'), dict(t='الاسم', k='name'),
      dict(t='شعير', k='worker'), dict(t='سرس', k='barley'), dict(t='أبيض', k='white'),
      dict(t='القيمة جنيه', k='val', n=1), dict(t='المصدر', k='ref', ref=1)]
write('expenses.html', page(
    'مصروفات الفترة الثانية',
    'المصادر: 2الحركة ف2 · شيت مصروفات ومسير مرتبات · الفترة الثانية فقط',
    [('exsum', f'ملخص المصروفات ({len(expsum)} بند)',
      f'إجمالي الملخص {money(expsum_total)} جنيه من شيت مصروفات الفترة الثانية.',
      EC, expsum, 50),
     ('exp', f'تفاصيل المصروفات ({len(exp):,} سطر)',
      f'إجمالي التفاصيل {money(exp_total)} جنيه. تحفظ: الملخص الجانبي يظهر أجور 1,280,373 مقابل 421,185 في التفاصيل، وإكراميات 340,268 مقابل 100,449، والفرق محتاج مراجعة.',
      XD, exp, 100),
     ('pay', f'مسير مرتبات ({len(pay):,} سطر)',
      f'إجمالي المسير {money(pay_total)} جنيه من مسير مرتبات الفترة الثانية.',
      PC, pay, 100)]))

# ============ الصفحة الرئيسية ============
cards = [
    ('شراء', 'purchases.html', f'{len(buy):,}', f'{money(buy_qty, 1)} طن', 'مشتريات وموردين · ميزان الموردين'),
    ('بيع', 'sales.html', f'{len(sell):,}', f'{money(sell_qty, 1)} طن', 'مبيعات وعملاء · ميزان العملاء'),
    ('ضرب وناولون', 'milling-freight.html', f'{len(mill) + len(freight):,}', f'{money(sum(x["val"] for x in freight if x["sid"] != 117) + 8445400 + 683598.5)} جنيه', 'حساب المضرب · الخزينة ف2'),
    ('مصروفات', 'expenses.html', f'{len(exp) + len(pay):,}', f'{money(exp_total + pay_total)} جنيه', 'شيت المصروفات · مسير المرتبات'),
]
h = [f'<!DOCTYPE html><html lang="ar" dir="rtl"><head><meta charset="utf-8">',
     '<meta name="viewport" content="width=device-width,initial-scale=1">',
     '<title>الفترة - 2 · مضرب الدمياطي استار</title><style>', CSS, '</style></head><body><div class="wrap">',
     '<header><h1>الفترة - 2</h1><div class="sub">مضرب الدمياطي استار · تفاصيل الفترة الثانية من ملفات الحركة الأصلية · كل رقم له مرجع لملف وشيت وسطر</div></header>',
     '<div class="links">']
for name, href, n, tot, src in cards:
    h.append(f'<a class="linkcard" href="{href}"><b>{name}</b><div class="sub">{src}</div>'
             f'<div style="margin-top:10px">{n} سطر · {tot}</div></a>')
# ===== تقارير ملخص الفترة الثانية فقط =====
freight_total = sum(x['val'] for x in freight if x['sid'] != 117)
buy_ret_total = sum(x['val'] for x in buy_ret)
sell_ret_total = sum(x['val'] for x in sell_ret)
summary = [
 ('ملخص الشراء', [
   ('قيمة مشتريات الفترة الثانية', f'{money(buy_val)} جنيه', 'purchases.html', 'تفاصيل الشراء'),
   ('وزن المشتريات', f'{money(buy_qty, 2)} طن', 'purchases.html', 'تفاصيل الشراء'),
   ('عدد السطور', f'{len(buy):,} سطر', 'purchases.html', 'تفاصيل الشراء'),
   ('مرتجع الشراء', f'{len(buy_ret)} سطر · {money(buy_ret_total)} جنيه', 'purchases.html', 'تفاصيل الشراء'),
   ('أعلى مورد قيمة', f'{sup_rows[0]["name"]} · {money(sup_rows[0]["val"])} جنيه', 'purchases.html', 'أعلى الموردين'),
 ]),
 ('ملخص البيع', [
   ('قيمة مبيعات الفترة الثانية', f'{money(sell_val)} جنيه', 'sales.html', 'تفاصيل البيع'),
   ('وزن المبيعات', f'{money(sell_qty, 2)} طن', 'sales.html', 'تفاصيل البيع'),
   ('عدد السطور', f'{len(sell):,} سطر', 'sales.html', 'تفاصيل البيع'),
   ('مرتجع البيع', f'{len(sell_ret)} سطر · {money(sell_ret_total)} جنيه', 'sales.html', 'تفاصيل البيع'),
   ('أعلى عميل قيمة', f'{cus_rows[0]["name"]} · {money(cus_rows[0]["val"])} جنيه', 'sales.html', 'أعلى العملاء'),
 ]),
 ('ملخص الضرب والناولون', [
   ('أجرة المضرب من حساب المضرب', '8,445,400.00 جنيه', 'milling-freight.html', 'حساب المضرب'),
   ('أجرة مضرب 650 للطن', '683,598.50 جنيه', 'milling-freight.html', 'الخزينة Sheet1'),
   ('نولون وبنود النقل في ف2', f'{money(freight_total)} جنيه', 'milling-freight.html', 'بنود ف2'),
   ('الإجمالي', f'{money(freight_total + 8445400 + 683598.5)} جنيه', 'milling-freight.html', 'كل البنود'),
 ]),
 ('ملخص المصروفات', [
   ('ملخص المصروفات', f'{money(expsum_total)} جنيه', 'expenses.html', 'ملخص المصروفات'),
   ('تفاصيل المصروفات', f'{money(exp_total)} جنيه', 'expenses.html', 'تفاصيل المصروفات'),
   ('مسير المرتبات', f'{money(pay_total)} جنيه', 'expenses.html', 'مسير المرتبات'),
   ('إجمالي المصروفات والمرتبات', f'{money(exp_total + pay_total)} جنيه', 'expenses.html', 'كل البنود'),
 ]),
]
h.append('<div class="sumhdr">تقارير ملخص الفترة الثانية</div>')
for title, rows in summary:
    h.append(f'<h2>{title}</h2>')
    h.append('<div class="tbox"><table><thead><tr><th>البند</th><th class="n">القيمة</th><th>المرجع</th></tr></thead><tbody>')
    for label, value, href, src in rows:
        last = label.startswith('إجمالي') or label.startswith('الإجمالي')
        cls = ' class="tot"' if last else ''
        h.append(f'<tr{cls}><td>{label}</td><td class="n">{value}</td>'
                 f'<td><a class="ref" target="_blank" href="{href}">{src}</a></td></tr>')
    h.append('</tbody></table></div>')
h.append('<div class="note">كل الأرقام والتفاصيل في هذه الصفحة خاصة بالفترة الثانية فقط، وكل بند يفتح التفاصيل والمرجع الأصلي للملف والشيت والسطر.</div>')
h.append('</div><div class="note">المصادر الأصلية للفترة الثانية فقط: 2الحركة (بدون باسورد).xlsm · حركه تعديل (بدون باسورد).xlsx · الخزينة.xlsx. كل صفحة تعرض تفاصيل الفترة الثانية بمرجع كل سطر.</div>')
h.append('</div></body></html>')
write('index.html', '\n'.join(h))
print('تم')
