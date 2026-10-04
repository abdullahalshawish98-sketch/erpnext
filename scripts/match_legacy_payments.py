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
  تاريخ المعاملة يُستخدم فقط لترجيح الأقرب عند تعدد المرشحين، لا للاستبعاد —
  لأن الدفعة قد تأتي بعد الطلب بأسابيع (بعد وصول الشحنة).
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


def names_match(a, b):
    ta, tb = name_tokens(a), name_tokens(b)
    if not ta or not tb:
        return False
    return bool(ta & tb)


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
    agent_raw = payment["agent"]
    mapped = AGENT_MAP.get(agent_raw, agent_raw if agent_raw else None)
    pool = sos
    if mapped:
        pool = [s for s in sos if s["sales_partner"] == mapped]
    return [s for s in pool if names_match(payment["customer"], s["customer_name"])]


def main():
    sos = load_sales_orders()
    payments = load_payments()
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
            cands = candidates_for(p1, sos) or candidates_for(p2, sos)
            exact = [s for s in cands if abs(s["grand_total"] - total) <= 1]
            if len(exact) == 1:
                results[p1["row"]] = (exact[0]["name"], f"دفعتان مجموعهما لطلب واحد (مع صف {p2['row']})")
                results[p2["row"]] = (exact[0]["name"], f"دفعتان مجموعهما لطلب واحد (مع صف {p1['row']})")
                used_pairs.add(p1["row"])
                used_pairs.add(p2["row"])
                break

    matched = sum(1 for p in payments if p["row"] in results)
    print(f"تمت المطابقة: {matched} من {len(payments)}")

    # كتابة النتيجة في نسخة من الإكسل
    wb = openpyxl.load_workbook(XLSX_IN)
    ws = wb["Sheet1"]
    ws.cell(row=1, column=5, value="رقم طلب البيع المطابَق")
    ws.cell(row=1, column=6, value="طريقة المطابقة")
    for p in payments:
        r, m = results.get(p["row"], (None, None))
        ws.cell(row=p["row"], column=5, value=r or "")
        ws.cell(row=p["row"], column=6, value=m or "لم تُطابَق")
    wb.save(XLSX_OUT)
    print(f"حُفظ الملف في {XLSX_OUT}")


if __name__ == "__main__":
    main()
