"""
يقارن تصدير حركات شي ان الفعلية ببطاقة الهدية في ERPNext. كل عملية استرجاع
تُنسَب لأقدم عملية شراء لم تُستهلَك بعد بنفس الحساب الفرعي (FIFO لكل حساب
على حدة) - عادة ما يأتي الاسترجاع بعد الشراء بحوالي ساعة. عمليات الشراء
بلا استرجاع تبقى بقيمتها الكاملة.
الاستخدام: python scripts/reconcile_gift_card_movements.py <مسار ملف حركات شي ان xlsx>
"""
import sys
import openpyxl, json, re
from collections import deque, defaultdict

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


def net_fifo_per_account(trs):
    """FIFO منفصل لكل حساب فرعي: كل استرجاع يُنسَب لأقدم شراء لم يُستهلَك
    بعد بنفس الحساب تحديداً، لا بأي حساب آخر."""
    events = sorted((t for t in trs if t[7] != 'Activated'), key=lambda t: (t[4], t[5]))
    open_by_acc = defaultdict(deque)
    resolved = []
    for t in events:
        acc, typ = t[6], t[7]
        if typ == 'Purchase':
            open_by_acc[acc].append({'amount': -t[8], 'date': t[4], 'time': t[5]})
        elif typ == 'Refund':
            q = open_by_acc[acc]
            if q:
                p = q.popleft()
                p['amount'] = round(p['amount'] - t[8], 2)
                resolved.append(p)
    for q in open_by_acc.values():
        resolved.extend(q)
    resolved.sort(key=lambda p: (p['date'], p['time']))
    return resolved


grand_unexplained = 0
for cid in shein_cards:
    name = card_by_serial.get(cid)
    erp_rows = sorted([p for p in pis if p.get('gift_card') == name and p.get('docstatus') == 1],
                       key=lambda p: (p.get('posting_date', ''), p.get('creation', '')))
    erp_pool = [(round(p.get('total') or 0, 2), p.get('name')) for p in erp_rows]

    nets = net_fifo_per_account(shein_cards[cid])
    print('===', cid, '(', len(nets), 'عملية صافية مقابل', len(erp_pool), 'فاتورة ) ===')

    unexplained = []
    for n in nets:
        match = None
        for ev, en in erp_pool:
            if abs(ev - n['amount']) < 0.05:
                match = en
                erp_pool.remove((ev, en))
                break
        if match:
            status = f'مطابق {match}'
        elif abs(n['amount']) < 0.05:
            status = 'استرجاع كامل (صافي صفر، لا فاتورة متوقَّعة)'
        else:
            status = 'غير مفسَّر'
            unexplained.append(n)
        print(f"  {n['date']} {n['time']}  صافي={n['amount']:8.2f}$  -> {status}")

    total_unexp = round(sum(n['amount'] for n in unexplained), 2)
    print(f'  >>> غير المفسَّر فعلياً: {total_unexp}$  ({len(unexplained)} عملية)')
    if erp_pool:
        print('  فواتير ERPNext متبقية (غير مُستخدَمة إطلاقاً):', erp_pool)
    grand_unexplained += total_unexp
    print()

print('==================================================')
print('الإجمالي العام غير المفسَّر:', round(grand_unexplained, 2))
