"""
مطابقة دفعات قديمة (قبل استخدام نظام الدفعات في ERPNext، مسجّلة في ملف إكسل خارجي)
بطلبات البيع الموجودة في ERPNext. قراءة فقط — لا تعديل على ERPNext، فقط يضيف
عموداً جديداً لملف الإكسل المحلي (نسخة محلية).

معايير المطابقة (بالترتيب):
  1. نفس الوكيل (بعد تطبيع الاسم) إن كان معروفاً في ملف الإكسل.
  2. تشابه اسم العميل (تطابق كلمة مشتركة بين الاسمين بعد حذف الأرقام والمسافات الزائدة).
  3. قيمة الدفعة = grand_total لطلب بيع واحد (بسماحية ±1).
  4. عند عدم وجود مطابقة مفردة: دفعتان لنفس العميل والوكيل مجموعهما = grand_total
     لطلب بيع واحد (بسماحية ±1).
  قيد التاريخ (بطلب المستخدم): لا يمكن أن تكون الدفعة أسبق من تاريخ طلب البيع
  بأكثر من 3 أيام (الدفعة قد تتأخر عن الطلب أسابيع، لكنها لا تتقدّمه إلا بأيام
  قليلة فقط) — فيُستبعد أي مرشح يخالف هذا القيد قبل أي معيار آخر.
"""
import json
import re
import openpyxl
from datetime import datetime

XLSX_IN = "/root/.claude/uploads/07a21562-e623-57de-a7c6-ab6649a6aa2d/c04a8e2c-_____.xlsx"
XLSX_OUT = "/home/user/erpnext/data/legacy_payments_matched.xlsx"

AGENT_MAP = {
    "سلتك": "متجر سلتك",
    "سيزون": "سيزون ستور",
    "عالم نون": "عالم نون",
    "وصلها": "متجر وصلها",
    "الاقصى": None,
    "غير معروف": None,
    "": None,
}


def norm_name(s):
    s = (s or "").strip()
    s = re.sub(r"\d+$", "", s)  # حذف أرقام تمييزية لاحقة مثل "فاطمة5"
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def name_tokens(s):
    return {t for t in norm_name(s).split(" ") if len(t) >= 2}


def build_stopwords(sos, max_df=8):
    """كلمات شائعة جداً بين أسماء العملاء (مثل 'متجر' أو 'محمد') لا تكفي وحدها
    كإشارة تطابق — تُستبعد من الاعتماد عليها بمفردها."""
    from collections import Counter
    names = {norm_name(s["customer_name"]) for s in sos}
    df = Counter()
    for n in names:
        for t in {t for t in n.split(" ") if len(t) >= 2}:
            df[t] += 1
    return {t for t, c in df.items() if c > max_df}


STOPWORDS = frozenset()  # تُهيَّأ في main() من بيانات العملاء الفعلية


def names_match(a, b, require_distinctive=True):
    """require_distinctive=True: كلمة مشتركة غير شائعة مطلوبة (الاسم هو الإشارة
    الوحيدة). False: أي كلمة مشتركة تكفي (يُستخدم حين يوجد تأكيد إضافي من
    الوكيل+القيمة+التاريخ معاً، كما طلب المستخدم)."""
    ta, tb = name_tokens(a), name_tokens(b)
    if not ta or not tb:
        return False
    shared = ta & tb
    if not shared:
        return False
    return bool(shared - STOPWORDS) if require_distinctive else True


def date_ok(payment_date, so_date_str):
    """الدفعة لا تتقدّم على طلب البيع بأكثر من 3 أيام (قد تتأخر عنه بلا حد)."""
    if not payment_date or not so_date_str:
        return False
    so_date = datetime.strptime(so_date_str, "%Y-%m-%d").date()
    return (so_date - payment_date).days <= 3


def load_sales_orders():
    with open("data/Sales_Order.json", encoding="utf-8") as f:
        sos = json.load(f)
    out = []
    for s in sos:
        if s.get("docstatus") != 1:
            continue
        td = s.get("transaction_date")
        out.append({
            "name": s.get("name"),
            "sales_partner": (s.get("sales_partner") or "").strip(),
            "customer_name": s.get("customer_name") or s.get("customer") or "",
            "transaction_date": td,
            "grand_total": round(float(s.get("base_grand_total") or s.get("grand_total") or 0), 2),
        })
    return out


def load_payments():
    wb = openpyxl.load_workbook(XLSX_IN, data_only=True)
    ws = wb["Sheet1"]
    rows = []
    for i, r in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        agent, customer, amount, date = r[0], r[1], r[2], r[3]
        if customer is None and amount is None:
            continue
        rows.append({
            "row": i, "agent": (agent or "").strip(), "customer": customer or "",
            "amount": round(float(amount), 2) if amount is not None else None,
            "date": date.date() if isinstance(date, datetime) else date,
        })
    return rows


def candidates_for(payment, sos):
    """جُرِّب السماح بكلمة اسم شائعة عند تطابق الوكيل (على افتراض أن
    الوكيل+القيمة+التاريخ يعوّضون ضعف إشارة الاسم) فتبيّن أنه غير آمن: وكيل
    واحد يخدم مئات العملاء، فكلمة مثل "متجر" أو "محمد" تطابق عشرات منهم —
    فعاد "متجر الحناشي" يطابق "متجر سلسبيل" و"امنة محمد" يطابق "معاوية محمد"
    زوراً. لذا يُطلب دوماً اسم مميّز (غير شائع) بصرف النظر عن تطابق الوكيل."""
    agent_raw = payment["agent"]
    mapped = AGENT_MAP.get(agent_raw, agent_raw if agent_raw else None)
    pool = sos
    if mapped:
        pool = [s for s in sos if s["sales_partner"] == mapped]
    return [s for s in pool if names_match(payment["customer"], s["customer_name"])
            and date_ok(payment["date"], s["transaction_date"])]


def main():
    global STOPWORDS
    sos = load_sales_orders()
    STOPWORDS = build_stopwords(sos)
    payments = load_payments()
    print(f"كلمات شائعة مُستبعدة من الاعتماد عليها وحدها: {len(STOPWORDS)}")
    print(f"طلبات البيع المرشَّحة (docstatus=1): {len(sos)}")
    print(f"صفوف الدفعات: {len(payments)}")

    results = {}  # row -> (so_name_or_names, method)

    # المرحلة 1: مطابقة مفردة (دفعة واحدة = طلب واحد)
    for p in payments:
        cands = candidates_for(p, sos)
        exact = [s for s in cands if p["amount"] is not None and abs(s["grand_total"] - p["amount"]) <= 1]
        if len(exact) == 1:
            results[p["row"]] = (exact[0]["name"], "مطابقة مباشرة (قيمة+اسم+وكيل)")
        elif len(exact) > 1:
            # رجّح الأقرب تاريخياً
            def key(s):
                if not s["transaction_date"] or not p["date"]:
                    return 9999
                d1 = datetime.strptime(s["transaction_date"], "%Y-%m-%d").date()
                return abs((p["date"] - d1).days)
            exact.sort(key=key)
            results[p["row"]] = (exact[0]["name"], f"مطابقة متعددة المرشحين ({len(exact)}) - اختير الأقرب تاريخياً")

    # المرحلة 2: دفعتان تجمعان لطلب واحد (لنفس العميل/الوكيل، غير مطابَقتين في المرحلة 1)
    unmatched = [p for p in payments if p["row"] not in results and p["amount"] is not None]
    used_pairs = set()
    for i, p1 in enumerate(unmatched):
        for p2 in unmatched[i+1:]:
            if p1["row"] in used_pairs or p2["row"] in used_pairs:
                continue
            if not names_match(p1["customer"], p2["customer"]):
                continue
            a1 = AGENT_MAP.get(p1["agent"], p1["agent"] or None)
            a2 = AGENT_MAP.get(p2["agent"], p2["agent"] or None)
            if a1 and a2 and a1 != a2:
                continue
            total = round(p1["amount"] + p2["amount"], 2)
            earliest = min(d for d in (p1["date"], p2["date"]) if d) if (p1["date"] or p2["date"]) else None
            mapped1 = AGENT_MAP.get(p1["agent"], p1["agent"] or None)
            pool = [s for s in sos if not mapped1 or s["sales_partner"] == mapped1]
            cands = [s for s in pool if names_match(p1["customer"], s["customer_name"])
                     and date_ok(earliest, s["transaction_date"])]
            exact = [s for s in cands if abs(s["grand_total"] - total) <= 1]
            if len(exact) == 1:
                results[p1["row"]] = (exact[0]["name"], f"دفعتان مجموعهما لطلب واحد (مع صف {p2['row']})")
                results[p2["row"]] = (exact[0]["name"], f"دفعتان مجموعهما لطلب واحد (مع صف {p1['row']})")
                used_pairs.add(p1["row"])
                used_pairs.add(p2["row"])
                break

    # المرحلة 3: مطابقة مفردة بتجاهل الوكيل (اسم+قيمة فقط) لمن تبقّى بلا نتيجة
    unmatched2 = [p for p in payments if p["row"] not in results and p["amount"] is not None]
    for p in unmatched2:
        cands = [s for s in sos if names_match(p["customer"], s["customer_name"])
                 and date_ok(p["date"], s["transaction_date"])
                 and abs(s["grand_total"] - p["amount"]) <= 1]
        if len(cands) == 1:
            results[p["row"]] = (cands[0]["name"], "مطابقة بتجاهل الوكيل (اسم+قيمة فقط) - يحتاج تأكيد")

    # المرحلة 4: دفعتان (بتجاهل الوكيل) تجمعان لطلب واحد
    unmatched3 = [p for p in payments if p["row"] not in results and p["amount"] is not None]
    used_pairs2 = set()
    for i, p1 in enumerate(unmatched3):
        for p2 in unmatched3[i+1:]:
            if p1["row"] in used_pairs2 or p2["row"] in used_pairs2:
                continue
            if not names_match(p1["customer"], p2["customer"]):
                continue
            total = round(p1["amount"] + p2["amount"], 2)
            earliest = min(d for d in (p1["date"], p2["date"]) if d) if (p1["date"] or p2["date"]) else None
            cands = [s for s in sos if names_match(p1["customer"], s["customer_name"])
                     and date_ok(earliest, s["transaction_date"])
                     and abs(s["grand_total"] - total) <= 1]
            if len(cands) == 1:
                results[p1["row"]] = (cands[0]["name"], f"دفعتان (بتجاهل الوكيل) مع صف {p2['row']} - يحتاج تأكيد")
                results[p2["row"]] = (cands[0]["name"], f"دفعتان (بتجاهل الوكيل) مع صف {p1['row']} - يحتاج تأكيد")
                used_pairs2.add(p1["row"])
                used_pairs2.add(p2["row"])
                break

    # تقرير للمرشحين المتعددين (بلا بت تلقائي) لمن تبقّى بلا نتيجة بعد كل المراحل
    still_unmatched = [p for p in payments if p["row"] not in results and p["amount"] is not None]
    multi_candidates = {}
    for p in still_unmatched:
        cands = [s for s in sos if names_match(p["customer"], s["customer_name"])
                 and date_ok(p["date"], s["transaction_date"])
                 and abs(s["grand_total"] - p["amount"]) <= 1]
        if cands:
            multi_candidates[p["row"]] = [c["name"] for c in cands]

    matched = sum(1 for p in payments if p["row"] in results)
    print(f"تمت المطابقة: {matched} من {len(payments)}")
    print(f"مرشحون متعددون بلا بتّ (اسم+قيمة): {len(multi_candidates)}")
    print(f"بلا أي مرشح إطلاقاً: {len(payments) - matched - len(multi_candidates)}")

    # كتابة النتيجة في نسخة من الإكسل
    wb = openpyxl.load_workbook(XLSX_IN)
    ws = wb["Sheet1"]
    ws.cell(row=1, column=5, value="رقم طلب البيع المطابَق")
    ws.cell(row=1, column=6, value="طريقة المطابقة")
    for p in payments:
        if p["row"] in results:
            r, m = results[p["row"]]
        elif p["row"] in multi_candidates:
            r, m = " / ".join(multi_candidates[p["row"]]), f"مرشحون متعددون ({len(multi_candidates[p['row']])}) بلا بتّ - يحتاج اختياراً يدوياً"
        else:
            r, m = "", "لا يوجد أي مرشح (لا تشابه اسم ولا قيمة مطابِقة)"
        ws.cell(row=p["row"], column=5, value=r)
        ws.cell(row=p["row"], column=6, value=m)
    wb.save(XLSX_OUT)
    print(f"حُفظ الملف في {XLSX_OUT}")


if __name__ == "__main__":
    main()
