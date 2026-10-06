"""
يقارن تصدير حركات شي ان الفعلية (عمود Account يحدد الحساب الفرعي) بفواتير
الشراء المرتبطة بنفس بطاقة الهدية في ERPNext، مع نسب كل استرجاع لآخر عملية
شراء غير مُستهلَكة بنفس الحساب الفرعي (لا بأي عملية شراء أقدم).
الاستخدام: python scripts/reconcile_gift_card_movements.py <مسار ملف حركات شي ان xlsx>
"""
import sys
import openpyxl, json, re
from collections import defaultdict

shein_file = sys.argv[1] if len(sys.argv) > 1 else 'data/shein_all_movements.xlsx'
wb = openpyxl.load_workbook(shein_file, data_only=True)
ws = wb['All Movements']
rows = list(ws.iter_rows(values_only=True))[1:]

shein_cards = {}
for r in rows:
    cid = str(r[2])
    shein_cards.setdefault(cid, []).append(r)

with open('data/Gift_Card.json', encoding='utf-8') as f:
    cards = json.load(f)
with open('data/Purchase_Invoice.json', encoding='utf-8') as f:
    pis = json.load(f)


def base_serial(name):
    m = re.search(r'(\d{10,})', name)
    return m.group(1) if m else None


card_by_serial = {}
for c in cards:
    bs = base_serial(c['name'])
    if bs:
        card_by_serial[bs] = c['name']


def net_per_purchase(trs):
    """لكل حساب فرعي: كل استرجاع يُطرَح من آخر عملية شراء غير مُستهلَكة قبله بنفس الحساب."""
    by_acc_purchases = defaultdict(list)  # acc -> [ {amount, date, time, consumed_refund} ]
    events = [t for t in trs if t[7] != 'Activated']
    for t in events:
        acc, typ, amt = t[6], t[7], t[8]
        if typ == 'Purchase':
            by_acc_purchases[acc].append({'amount': -amt, 'date': t[4], 'time': t[5]})
        elif typ == 'Refund':
            lst = by_acc_purchases[acc]
            if lst:
                lst[-1]['amount'] -= amt  # يُطرَح من آخر شراء بنفس الحساب
            else:
                # استرجاع بلا شراء سابق بنفس الحساب (نادر) - نتجاهله كحدث منفصل
                pass
    out = []
    for acc, lst in by_acc_purchases.items():
        for p in lst:
            out.append({'account': acc, 'date': p['date'], 'time': p['time'], 'net': round(p['amount'], 2)})
    out.sort(key=lambda x: (x['date'], x['time']))
    return out


grand_unexplained = 0
for cid in shein_cards:
    name = card_by_serial.get(cid)
    erp_rows = sorted([p for p in pis if p.get('gift_card') == name and p.get('docstatus') == 1],
                       key=lambda p: (p.get('posting_date', ''), p.get('creation', '')))
    erp_pool = [(round(p.get('total') or 0, 2), p.get('name')) for p in erp_rows]

    nets = net_per_purchase(shein_cards[cid])
    print('===', cid, '===')

    unexplained = []
    for n in nets:
        match = None
        for ev, en in erp_pool:
            if abs(ev - n['net']) < 0.05:
                match = en
                erp_pool.remove((ev, en))
                break
        status = f'مطابق {match}' if match else 'غير مفسَّر'
        print(f"  {n['date']} {n['time']}  حساب {n['account']:4s}  صافي={n['net']:8.2f}$  -> {status}")
        if not match:
            unexplained.append(n)

    total_unexp = round(sum(n['net'] for n in unexplained), 2)
    print(f'  >>> غير المفسَّر فعلياً: {total_unexp}$  ({len(unexplained)} عملية)')
    if erp_pool:
        print('  فواتير ERPNext متبقية (غير مُستخدَمة إطلاقاً):', erp_pool)
    grand_unexplained += total_unexp
    print()

print('==================================================')
print('الإجمالي العام غير المفسَّر:', round(grand_unexplained, 2))
