"""
يقارن تصدير حركات شي ان الفعلية ببطاقة الهدية في ERPNext، حساباً فرعياً
حساباً فرعياً. كل استرجاع يُقرَن بعملية الشراء الأقرب زمنياً لفارق 60 دقيقة
بالضبط (قاعدة "الاسترجاع يرجع بعد ساعة تقريباً من الدفع")، وليس بترتيب
تسلسلي بسيط. إن تعادل مجموع استرجاعات دفعة كاملة مع مجموع مشترياتها،
تُعتبر الدفعة ملغاة بالكامل (صافٍ صفر، لا فاتورة متوقَّعة).
الاستخدام: python scripts/reconcile_gift_card_movements.py <مسار ملف حركات شي ان xlsx>
"""
import sys
import openpyxl, json, re
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


def reconcile_account(purchases, refunds):
    """مطابقة جشعة: كل زوج (شراء، استرجاع) مرشَّح حسب أقرب فارق لـ60 دقيقة،
    تُرتَّب كل الأزواج الممكنة تصاعدياً حسب |الفارق - 60min| وتُؤخَذ الأفضل أولاً."""
    # إلغاء كامل للجلسة: إن تساوى مجموع كل مشترياتها مع مجموع كل استرجاعاتها،
    # فالأرجح أنها كلها أُرجعت (لا حاجة لمحاولة إقران فردي غامض)
    if purchases and refunds and abs(sum(p['amount'] for p in purchases) - sum(r['amount'] for r in refunds)) < 0.05:
        results = [{**p, 'net': 0.0, 'note': 'جلسة أُلغيت بالكامل (مجموع الشراء = مجموع الاسترجاع)'} for p in purchases]
        return results, []

    candidates = []
    for pi, p in enumerate(purchases):
        for ri, r in enumerate(refunds):
            delta_min = (dt(r['date'], r['time']) - dt(p['date'], p['time'])).total_seconds() / 60
            if 10 <= delta_min <= 180:  # نافذة معقولة حول الساعة
                candidates.append((abs(delta_min - 60), pi, ri))
    candidates.sort(key=lambda x: x[0])

    used_p, used_r = set(), set()
    pair_of = {}
    for _, pi, ri in candidates:
        if pi in used_p or ri in used_r:
            continue
        used_p.add(pi)
        used_r.add(ri)
        pair_of[pi] = ri

    # دفعة كاملة الإلغاء: كل المشتريات غير المُقرَنة + كل الاسترجاعات غير المُقرَنة
    # ضمن نفس الحساب، إن تساوى مجموعهما، تُعتبر كلها صافي صفر
    unpaired_p_idx = [i for i in range(len(purchases)) if i not in used_p]
    unpaired_r_idx = [i for i in range(len(refunds)) if i not in used_r]
    full_cancel = set()
    if unpaired_p_idx and unpaired_r_idx:
        sp = sum(purchases[i]['amount'] for i in unpaired_p_idx)
        sr = sum(refunds[i]['amount'] for i in unpaired_r_idx)
        if abs(sp - sr) < 0.05:
            full_cancel = set(unpaired_p_idx)

    results = []
    for pi, p in enumerate(purchases):
        if pi in pair_of:
            net = round(p['amount'] - refunds[pair_of[pi]]['amount'], 2)
            results.append({**p, 'net': net, 'note': f"مقرون باسترجاع {refunds[pair_of[pi]]['time']}"})
        elif pi in full_cancel:
            results.append({**p, 'net': 0.0, 'note': 'ضمن دفعة أُلغيت بالكامل'})
        else:
            results.append({**p, 'net': round(p['amount'], 2), 'note': 'بلا استرجاع'})
    leftover_refunds = [refunds[i] for i in unpaired_r_idx if i not in full_cancel]
    return results, leftover_refunds


grand_unexplained = 0
for cid in shein_cards:
    name = card_by_serial.get(cid)
    erp_rows = sorted([p for p in pis if p.get('gift_card') == name and p.get('docstatus') == 1],
                       key=lambda p: (p.get('posting_date', ''), p.get('creation', '')))
    erp_pool = [(round(p.get('total') or 0, 2), p.get('name')) for p in erp_rows]

    accounts = sorted({t[6] for t in shein_cards[cid] if t[7] != 'Activated'})
    all_results = []
    for acc in accounts:
        events = sorted((t for t in shein_cards[cid] if t[7] != 'Activated' and t[6] == acc), key=lambda t: (t[4], t[5]))
        # تقسيم نشاط الحساب إلى جلسات منفصلة عند أي فجوة زمنية > 3 ساعات، لتفادي
        # خلط دفعتين مختلفتين (صباحية ومسائية مثلاً) تحت نفس رمز الحساب المختصر
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

        for sess in sessions:
            purchases = [{'amount': -t[8], 'date': t[4], 'time': t[5]} for t in sess if t[7] == 'Purchase']
            refunds = [{'amount': t[8], 'date': t[4], 'time': t[5]} for t in sess if t[7] == 'Refund']
            results, leftover = reconcile_account(purchases, refunds)
            all_results.extend(results)
            if leftover:
                print(f'  [تنبيه] حساب {acc} في البطاقة {cid}: استرجاعات بلا شراء مقابل واضح: {[r["amount"] for r in leftover]}')

    all_results.sort(key=lambda x: (x['date'], x['time']))
    print('===', cid, '===')
    unexplained = []
    for n in all_results:
        if abs(n['net']) < 0.05:
            print(f"  {n['date']} {n['time']}  صافي={n['net']:8.2f}$  -> {n['note']}")
            continue
        match = None
        for ev, en in erp_pool:
            if abs(ev - n['net']) < 0.05:
                match = en
                erp_pool.remove((ev, en))
                break
        status = f'مطابق {match}' if match else 'غير مفسَّر'
        print(f"  {n['date']} {n['time']}  صافي={n['net']:8.2f}$  ({n['note']})  -> {status}")
        if not match:
            unexplained.append(n)

    total_unexp = round(sum(n['net'] for n in unexplained), 2)
    print(f'  >>> غير مفسَّر: {total_unexp}$  ({len(unexplained)} عملية)')
    if erp_pool:
        print('  فواتير ERPNext متبقية (غير مُستخدَمة إطلاقاً):', erp_pool)
    grand_unexplained += total_unexp
    print()

print('==================================================')
print('الإجمالي العام غير المفسَّر:', round(grand_unexplained, 2))
