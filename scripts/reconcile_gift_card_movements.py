"""
يقارن تصدير حركات شي ان الفعلية ببطاقة الهدية في ERPNext، حساباً فرعياً
حساباً فرعياً. كل استرجاع يُقرَن بعملية الشراء الأقرب زمنياً لفارق 60 دقيقة
بالضبط (قاعدة "الاسترجاع يرجع بعد ساعة تقريباً من الدفع")، وليس بترتيب
تسلسلي بسيط. إن تعادل مجموع استرجاعات دفعة كاملة مع مجموع مشترياتها،
تُعتبر الدفعة ملغاة بالكامل (صافٍ صفر، لا فاتورة متوقَّعة).
يحفظ أيضاً كل العمليات المدفوعة وغير المُفسَّرة (تاريخ، وقت، حساب، مبلغ) في
ملف CSV لكل البطاقات مجتمعة.
الاستخدام: python scripts/reconcile_gift_card_movements.py <مسار ملف حركات شي ان xlsx> [مسار ملف الإخراج CSV]
"""
import sys
import openpyxl, json, re, csv
from datetime import datetime
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


def dt(date_, time_):
    return datetime.fromisoformat(f"{date_} {time_}")


def take(erp_by_value, value):
    v = round(value, 2)
    names = erp_by_value.get(v)
    if names:
        return names.pop()
    return None


def candidates(purchases, refunds, p_idx, r_idx, lo_min, hi_min, target_min):
    cands = []
    for pi in p_idx:
        for ri in r_idx:
            delta_min = (dt(refunds[ri]['date'], refunds[ri]['time']) - dt(purchases[pi]['date'], purchases[pi]['time'])).total_seconds() / 60
            if lo_min <= delta_min <= hi_min:
                cands.append((abs(delta_min - target_min), pi, ri))
    cands.sort(key=lambda x: x[0])
    return cands


def reconcile_session(purchases, refunds, erp_by_value):
    """مرحلة أولى لكل جلسة (نشاط متقارب زمنياً لنفس الحساب):
    1) إلغاء كامل للجلسة إن تساوى مجموع الشراء مع مجموع الاسترجاع (قطعة أُضيفت
       لتفادي رسوم التوصيل على الطلبات الأقل من 41$، أُلغيت خلال الساعة الأولى).
    2) لكل شراء: إن طابقت قيمته الكاملة فاتورة متاحة مباشرة، يُقبَل بلا استرجاع.
    3) إقران عادي خلال 10-180 دقيقة (الأقرب لـ60 دقيقة) — النمط الطبيعي.
    يُعيد: (النتائج، فهارس المشتريات غير المحسومة، فهارس الاسترجاعات غير المحسومة)
    مع المشتريات/الاسترجاعات الأصلية، لتُمرَّر لاحقاً لمرحلة الإقران الموسَّع
    على مستوى الحساب كله (قد تمتد عبر عدة جلسات)."""
    if purchases and refunds and abs(sum(p['amount'] for p in purchases) - sum(r['amount'] for r in refunds)) < 0.05:
        results = [{**p, 'net': 0.0, 'note': 'جلسة أُلغيت بالكامل (مجموع الشراء = مجموع الاسترجاع)', 'match': None} for p in purchases]
        return results, [], []

    all_p = list(range(len(purchases)))
    all_r = list(range(len(refunds)))
    used_p, used_r = set(), set()
    note_of = {}

    for pi in all_p:
        matched_name = take(erp_by_value, purchases[pi]['amount'])
        if matched_name:
            used_p.add(pi)
            note_of[pi] = ('مطابقة مباشرة', None, matched_name)

    remaining_p = [i for i in all_p if i not in used_p]
    for _, pi, ri in candidates(purchases, refunds, remaining_p, all_r, 10, 180, 60):
        if pi in used_p or ri in used_r:
            continue
        net = round(purchases[pi]['amount'] - refunds[ri]['amount'], 2)
        matched_name = take(erp_by_value, net)
        used_p.add(pi)
        used_r.add(ri)
        note_of[pi] = (f"مقرون باسترجاع {refunds[ri]['time']}", net, matched_name)

    results = []
    for pi, p in enumerate(purchases):
        if pi in note_of:
            note, net, matched_name = note_of[pi]
            results.append({**p, 'net': round(p['amount'], 2) if net is None else net, 'note': note, 'match': matched_name})
        else:
            results.append({**p, 'net': round(p['amount'], 2), 'note': 'بلا استرجاع (بعد)', 'match': None})
    unpaired_p_idx = [i for i in all_p if i not in used_p]
    unpaired_r_idx = [i for i in all_r if i not in used_r]
    return results, unpaired_p_idx, unpaired_r_idx


def wide_match(leftover_result_dicts, leftover_refunds, erp_by_value):
    """المرحلة الموسَّعة (حتى 24 ساعة)، على مستوى الحساب كله (قد تمتد عبر
    حدود الجلسة الضيقة) — تُقبَل فقط إن طابق الصافي فاتورة متاحة فعلاً، تفادياً
    لإفساد مشترياتٍ سليمة بتخمين زمني لا يؤكده أي دليل محاسبي. يُعدِّل
    leftover_result_dicts في مكانها (net/note/match) للعناصر التي أُقرِنت."""
    all_p = list(range(len(leftover_result_dicts)))
    all_r = list(range(len(leftover_refunds)))
    used_p, used_r = set(), set()
    for _, pi, ri in candidates(leftover_result_dicts, leftover_refunds, all_p, all_r, 180, 1440, 180):
        if pi in used_p or ri in used_r:
            continue
        net = round(leftover_result_dicts[pi]['net'] - leftover_refunds[ri]['amount'], 2)
        matched_name = take(erp_by_value, net)
        if matched_name:
            used_p.add(pi)
            used_r.add(ri)
            r = leftover_refunds[ri]
            leftover_result_dicts[pi]['net'] = net
            leftover_result_dicts[pi]['note'] = f"⚠ شاذة — مقرون باسترجاع متأخر {r['date']} {r['time']} (نسيان إلغاء محتمل)"
            leftover_result_dicts[pi]['match'] = matched_name


grand_unexplained = 0
all_unexplained_rows = []
for cid in shein_cards:
    name = card_by_serial.get(cid)
    erp_rows = sorted([p for p in pis if p.get('gift_card') == name and p.get('docstatus') == 1],
                       key=lambda p: (p.get('posting_date', ''), p.get('creation', '')))
    erp_by_value = defaultdict(list)
    for p in erp_rows:
        erp_by_value[round(p.get('total') or 0, 2)].append(p.get('name'))
    return_names = {p['name'] for p in erp_rows if p.get('is_return')}

    accounts = sorted({t[6] for t in shein_cards[cid] if t[7] != 'Activated'})
    all_results = []
    for acc in accounts:
        events = sorted((t for t in shein_cards[cid] if t[7] != 'Activated' and t[6] == acc), key=lambda t: (t[4], t[5]))
        # تقسيم نشاط الحساب إلى جلسات ضيقة عند أي فجوة زمنية > 3 ساعات، لحصر
        # كشف "الإلغاء الكامل" ضمن دفعة واحدة متقاربة فقط (لا تختلط دفعة صباحية
        # بأخرى مسائية تحت نفس رمز الحساب المختصر)
        sessions = []
        cur = []
        prev_dt = None
        for t in events:
            cur_dt = dt(t[4], t[5])
            if prev_dt and (cur_dt - prev_dt).total_seconds() > 3 * 3600:
                sessions.append(cur)
                cur = []
            cur.append(t)
            prev_dt = cur_dt
        if cur:
            sessions.append(cur)

        acc_leftover_purchases = []
        acc_leftover_refunds = []
        for sess in sessions:
            purchases = [{'amount': -t[8], 'date': t[4], 'time': t[5], 'account': acc} for t in sess if t[7] == 'Purchase']
            refunds = [{'amount': t[8], 'date': t[4], 'time': t[5], 'account': acc} for t in sess if t[7] == 'Refund']
            results, unpaired_p_idx, unpaired_r_idx = reconcile_session(purchases, refunds, erp_by_value)
            all_results.extend(results)
            acc_leftover_purchases.extend(results[i] for i in unpaired_p_idx)
            acc_leftover_refunds.extend(refunds[i] for i in unpaired_r_idx)

        # مرحلة الإقران الموسَّع (حتى 24 ساعة) على مستوى الحساب كله، قد تمتد
        # عبر حدود الجلسات الضيقة أعلاه
        wide_match(acc_leftover_purchases, acc_leftover_refunds, erp_by_value)

    all_results.sort(key=lambda x: (x['date'], x['time']))
    print('===', cid, '===')
    unexplained = []
    remaining_pool_names = {n for names in erp_by_value.values() for n in names}
    for n in all_results:
        if abs(n['net']) < 0.05:
            print(f"  {n['date']} {n['time']}  صافي={n['net']:8.2f}$  -> {n['note']}")
            continue
        match = n.get('match')
        if match:
            status = f'مطابق {match}'
        else:
            remaining_pool_names = {nm for names in erp_by_value.values() for nm in names}
            remaining_returns = return_names & remaining_pool_names
            if remaining_returns:
                status = f'📦 غير مفسَّر، لكن توجد فاتورة مرتجع غير مستخدَمة على نفس البطاقة ({", ".join(remaining_returns)}) — قد تكون حالة نفاذ مخزون'
            else:
                status = 'غير مفسَّر'
        print(f"  {n['date']} {n['time']}  صافي={n['net']:8.2f}$  ({n['note']})  -> {status}")
        if not match:
            unexplained.append(n)
            remaining_pool_names = {nm for names in erp_by_value.values() for nm in names}
            all_unexplained_rows.append({
                'card_id': cid, 'date': n['date'], 'time': n['time'],
                'account': n['account'], 'amount': n['net'],
                'possible_out_of_stock': bool(return_names & remaining_pool_names),
            })

    total_unexp = round(sum(n['net'] for n in unexplained), 2)
    print(f'  >>> غير مفسَّر: {total_unexp}$  ({len(unexplained)} عملية)')
    leftover_invoices = [(v, nm) for v, names in erp_by_value.items() for nm in names]
    if leftover_invoices:
        print('  فواتير ERPNext متبقية (غير مُستخدَمة إطلاقاً):', leftover_invoices)
    grand_unexplained += total_unexp
    print()

print('==================================================')
print('الإجمالي العام غير المفسَّر:', round(grand_unexplained, 2))

out_path = sys.argv[2] if len(sys.argv) > 2 else 'data/unexplained_shein_purchases.csv'
all_unexplained_rows.sort(key=lambda r: (r['card_id'], r['date'], r['time']))
with open(out_path, 'w', encoding='utf-8-sig', newline='') as f:
    w = csv.DictWriter(f, fieldnames=['card_id', 'date', 'time', 'account', 'amount', 'possible_out_of_stock'])
    w.writeheader()
    w.writerows(all_unexplained_rows)
print('تم حفظ العمليات غير المفسَّرة في:', out_path)
