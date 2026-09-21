#!/usr/bin/env python3
"""
كشف أخطاء الإدخال والقيم الشاذة والتطابق بين Sales Order / Purchase Invoice / Shipments
يعمل محلياً على ملفات data/*.json (ناتجة عن erpnext_explorer.py export --full).
لا يتصل بالـ API إطلاقاً. Standard library only.

الإخراج:
  data/flags.csv           كل الحالات المشبوهة (doc, doctype, agent, reason, severity, detail)
  data/agent_summary.csv   نسبة الطلبات المعلَّمة لكل وكيل
"""
import csv
import json
import statistics
from collections import Counter, defaultdict
from datetime import date, datetime
from pathlib import Path

import data_corrections
import exclusions

TODAY = date(2026, 9, 20)

# ---- thresholds (قابلة للتعديل) --------------------------------------
NO_PI_AFTER_DAYS = 3          # طلب بيع بدون فاتورة شراء بعد N أيام
SHIPPING_COST_PCT = 30        # سقف نسبة تكلفة الشحن من قيمة الطلب
FX_RATE_MAD_K = 3             # عامل الشذوذ للوسيط±MAD لسعر الصرف
ITEM_MAD_K = 3                # عامل الشذوذ للوسيط±MAD للسعر/الكمية لكل صنف/وكيل
MIN_GROUP_SIZE = 5             # أقل عدد عينات لحساب MAD لصنف/وكيل


def load(name):
    return json.load(open(Path("data") / f"{name}.json", encoding="utf-8"))


def parse_date(s):
    if not s:
        return None
    try:
        return date.fromisoformat(str(s)[:10])
    except ValueError:
        return None


def mad_bounds(values, k):
    """المدى الطبيعي = الوسيط ± k * 1.4826 * MAD"""
    vals = sorted(v for v in values if v is not None)
    if len(vals) < MIN_GROUP_SIZE:
        return None
    med = statistics.median(vals)
    mad = statistics.median([abs(v - med) for v in vals])
    if mad == 0:
        return None
    spread = k * 1.4826 * mad
    return med - spread, med + spread, med


class Flags:
    def __init__(self):
        self.rows = []

    def add(self, doc, doctype, agent, reason, severity, detail):
        self.rows.append({"doc": doc, "doctype": doctype, "agent": agent or "",
                           "reason": reason, "severity": severity, "detail": detail})

    def save(self, path):
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=["doc", "doctype", "agent", "reason", "severity", "detail"])
            w.writeheader()
            w.writerows(self.rows)


def main():
    so_all = load("Sales_Order")
    pi_all = load("Purchase_Invoice")
    sh_all = load("Shipments")
    data_corrections.apply(so_all, sh_all)
    so_all = exclusions.filter_out(so_all)

    # نتجاهل الطلبات/الفواتير/الشحنات الملغاة أو المسوَّدة تماماً من كل التحليل
    # (لا تُفحص، ولا تُدخَل في حسابات الوسيط/MAD المرجعية) — وليس فقط من التعليم.
    so = [r for r in so_all if r.get("docstatus") == 1]
    pi = [r for r in pi_all if r.get("docstatus") == 1]
    sh = [r for r in sh_all if r.get("docstatus") == 1]
    so_live_names = {r["name"] for r in so}

    so_by_name = {r["name"]: r for r in so}
    pi_by_name = {r["name"]: r for r in pi}

    pi_by_so = defaultdict(list)
    for r in pi:
        if r.get("sales_order") in so_live_names:
            pi_by_so[r["sales_order"]].append(r)

    sh_by_so = defaultdict(list)
    sh_by_pi = defaultdict(list)
    for r in sh:
        if r.get("sales_order") in so_live_names:
            sh_by_so[r["sales_order"]].append(r)
        if r.get("purchase_invoice"):
            sh_by_pi[r["purchase_invoice"]].append(r)

    # أسعار الصرف الشاذة تُحسب أولاً لأن فحوصاً أخرى (تكلفة الشحن) يجب أن
    # تتجاهل/تحذر عند الاعتماد على تحويل عملة غير موثوق. مبنية فقط من فواتير
    # شراء/شحن مرتبطة بطلبات بيع غير ملغاة، حتى لا تلوّث طلبات ملغاة المدى الطبيعي.
    pi_rates = [r.get("conversion_rate") for r in pi
                if r.get("conversion_rate") and r.get("sales_order") in so_live_names]
    pi_fx_bounds = mad_bounds(pi_rates, FX_RATE_MAD_K)
    bad_pi_fx = {r["name"] for r in pi if pi_fx_bounds and r.get("conversion_rate")
                 and not (pi_fx_bounds[0] <= r["conversion_rate"] <= pi_fx_bounds[1])}

    sh_rates = [r.get("exchange_rate") for r in sh
                if r.get("exchange_rate") and r.get("sales_order") in so_live_names]
    sh_fx_bounds = mad_bounds(sh_rates, FX_RATE_MAD_K)
    bad_sh_fx = {r["name"] for r in sh if sh_fx_bounds and r.get("exchange_rate")
                 and not (sh_fx_bounds[0] <= r["exchange_rate"] <= sh_fx_bounds[1])}

    flags = Flags()
    flagged_docs_by_agent = defaultdict(set)   # agent -> set(so_name) flagged
    total_docs_by_agent = Counter()

    for r in so:
        total_docs_by_agent[r.get("sales_partner") or "(بدون وكيل)"] += 1

    def flag_so(so_rec, reason, severity, detail):
        flags.add(so_rec["name"], "Sales Order", so_rec.get("sales_partner"), reason, severity, detail)
        flagged_docs_by_agent[so_rec.get("sales_partner") or "(بدون وكيل)"].add(so_rec["name"])

    # === فحص 1: طلب بيع بدون فاتورة شراء بعد X أيام ========================
    for r in so:
        if r["name"] in pi_by_so:
            continue
        d = parse_date(r.get("transaction_date"))
        if not d:
            continue
        age = (TODAY - d).days
        if age >= NO_PI_AFTER_DAYS:
            sev = "عالية" if age >= 14 else "متوسطة"
            flag_so(r, "طلب بيع بدون فاتورة شراء", sev,
                    f"عمره {age} يوم، تاريخ الطلب {r.get('transaction_date')}")

    # === فحص 2: إجمالي فاتورة الشراء لا يطابق الطلب =========================
    # النسبة الطبيعية هي هامش ربح (تكلفة الشراء أقل من سعر البيع)، لذا نستخدم
    # نسبة (مجموع فواتير الشراء / إجمالي الطلب) ونقارنها بمداها الطبيعي عبر
    # الوسيط±3MAD، مع تعليم صريح لأي حالة تتجاوز فيها التكلفة سعر البيع (خسارة مؤكدة).
    ratios = {}
    for so_name, pis in pi_by_so.items():
        r = so_by_name.get(so_name)
        if not r or not r.get("base_grand_total"):
            continue
        sum_pi = sum(p.get("base_grand_total") or 0 for p in pis)
        ratios[so_name] = (sum_pi, r["base_grand_total"], sum_pi / r["base_grand_total"], pis)
    rbounds = mad_bounds([v[2] for v in ratios.values()], FX_RATE_MAD_K)
    for so_name, (sum_pi, so_total, ratio, pis) in ratios.items():
        r = so_by_name[so_name]
        pi_list = ", ".join(p["name"] for p in pis)
        if ratio >= 1:
            flag_so(r, "إجمالي فاتورة الشراء يتجاوز إجمالي الطلب (خسارة مؤكدة)", "عالية",
                    f"تكلفة الشراء {sum_pi:,.2f} LYD >= سعر البيع {so_total:,.2f} LYD "
                    f"(نسبة التكلفة {ratio*100:.0f}%) — فواتير: {pi_list}")
        elif rbounds and not (rbounds[0] <= ratio <= rbounds[1]):
            sev = "عالية" if ratio > rbounds[1] else "متوسطة"
            note = "تكلفة مرتفعة بشكل غير معتاد (هامش ربح ضعيف)" if ratio > rbounds[1] else \
                   "تكلفة منخفضة بشكل غير معتاد (تحقق من اكتمال فواتير الشراء)"
            flag_so(r, "نسبة تكلفة الشراء إلى سعر البيع شاذة (وسيط±3MAD)", sev,
                    f"{note}: نسبة التكلفة {ratio*100:.0f}% خارج المدى الطبيعي "
                    f"[{rbounds[0]*100:.0f}%, {rbounds[1]*100:.0f}%] — فواتير: {pi_list}")

    # === فحص 3: الطلب مكتمل ولا توجد فاتورة شحن =============================
    for r in so:
        if r.get("status") == "Completed" and r["name"] not in sh_by_so:
            flag_so(r, "طلب مكتمل بدون فاتورة شحن", "عالية",
                    f"status=Completed، تاريخ الطلب {r.get('transaction_date')}")

    # === فحص 4: تكلفة الشحن أعلى من نسبة معقولة من قيمة الطلب ================
    # company_currency_total_cost = total_cost(USD) * exchange_rate، فهي بنفس
    # عملة الطلب (LYD) فعلاً — لكن هذا التحويل غير موثوق إن كان exchange_rate
    # نفسه شاذاً (انظر فحص أسعار الصرف)، فنستبعد تلك الحالات من هذا الفحص
    # ونتركها لفحص "سعر صرف شاذ" وحده حتى لا تُحسب نسبة مضلِّلة.
    for s in sh:
        if s["name"] in bad_sh_fx:
            continue
        so_name = s.get("sales_order")
        r = so_by_name.get(so_name)
        if not r or not r.get("base_grand_total"):
            continue
        cost = s.get("company_currency_total_cost") or 0  # LYD، نفس عملة base_grand_total
        ratio = cost / r["base_grand_total"] * 100
        if ratio > SHIPPING_COST_PCT:
            sev = "عالية" if ratio > 70 else "متوسطة"
            flag_so(r, "تكلفة الشحن مرتفعة نسبة لقيمة الطلب", sev,
                    f"فاتورة الشحن {s['name']}: تكلفة {cost:,.2f} LYD = {ratio:.0f}% من قيمة الطلب {r['base_grand_total']:,.2f} LYD")

    # === فحص 5: تسلسل تواريخ غير منطقي ======================================
    # تاريخ "وصول" الشحنة الفعلي هو arrival_date، وليس posting_date (تاريخ
    # قيد فاتورة الشحن إدارياً) — الاثنان مختلفان فعلياً بفارق أيام قد يصل
    # لأسابيع بسبب مدة الشحن، وهذا طبيعي. غير المنطقي هو أن "تصل" الشحنة
    # قبل أن "نشتريها" أصلاً (arrival_date قبل تاريخ فاتورة الشراء).
    for s in sh:
        pi_r = pi_by_name.get(s.get("purchase_invoice"))
        so_r = so_by_name.get(s.get("sales_order"))
        arrival = parse_date(s.get("arrival_date"))
        if pi_r and arrival and so_r:
            pi_date = parse_date(pi_r.get("posting_date"))
            if pi_date and arrival < pi_date:
                flag_so(so_r,
                        "تسلسل تواريخ غير منطقي: وصول الشحنة قبل تاريخ فاتورة الشراء", "عالية",
                        f"فاتورة الشحن {s['name']} وصلت بتاريخ {s.get('arrival_date')} قبل فاتورة الشراء "
                        f"{pi_r['name']} المؤرَّخة {pi_r.get('posting_date')}")
        if so_r:
            so_date = parse_date(so_r.get("transaction_date"))
            if pi_r and so_date:
                pi_date = parse_date(pi_r.get("posting_date"))
                if pi_date and pi_date < so_date:
                    flag_so(so_r, "تسلسل تواريخ غير منطقي: الشراء قبل تاريخ طلب البيع", "منخفضة",
                            f"فاتورة الشراء {pi_r['name']} بتاريخ {pi_r.get('posting_date')} قبل "
                            f"تاريخ الطلب {so_r.get('transaction_date')}")

    # === فحص 6: اختلاف العملة / سعر الصرف ===================================
    for r in pi:
        cr = r.get("conversion_rate")
        if cr == 1.0:
            so_r = so_by_name.get(r.get("sales_order"))
            if so_r:
                flag_so(so_r, "فاتورة شراء بدولار بسعر صرف = 1 (لم يُحدَّث)", "متوسطة",
                        f"{r['name']}: conversion_rate=1.0 مع عملة {r.get('currency')}")
        elif r["name"] in bad_pi_fx:
            so_r = so_by_name.get(r.get("sales_order"))
            if so_r:
                flag_so(so_r, "سعر صرف شاذ في فاتورة الشراء (وسيط±3MAD)", "عالية",
                        f"{r['name']}: rate={cr} خارج المدى الطبيعي [{pi_fx_bounds[0]:.2f}, {pi_fx_bounds[1]:.2f}] "
                        f"(الوسيط {pi_fx_bounds[2]:.2f})")
    for s in sh:
        if s["name"] in bad_sh_fx:
            so_r = so_by_name.get(s.get("sales_order"))
            if so_r:
                flag_so(so_r, "سعر صرف شاذ في فاتورة الشحن (وسيط±3MAD)", "عالية",
                        f"{s['name']}: exchange_rate={s.get('exchange_rate')} خارج المدى الطبيعي "
                        f"[{sh_fx_bounds[0]:.2f}, {sh_fx_bounds[1]:.2f}] (الوسيط {sh_fx_bounds[2]:.2f})")
    for r in so:
        if r.get("currency") != "LYD":
            flag_so(r, "عملة طلب البيع ليست عملة الشركة (LYD)", "منخفضة", f"currency={r.get('currency')}")

    # === فحص 7: كمية أو سعر صفر/سالب (بنود الطلب) ============================
    for r in so:
        for it in r.get("items", []):
            qty, rate = it.get("qty") or 0, it.get("rate") or 0
            if qty <= 0 or rate <= 0:
                flag_so(r, "كمية أو سعر صفر/سالب في بند الطلب", "عالية",
                        f"صنف {it.get('item_code')}: qty={qty}, rate={rate}")

    # === فحص 8: سعر بيع أقل من الشراء (لنفس الطلب/الفاتورة) =================
    for so_name, pis in pi_by_so.items():
        r = so_by_name.get(so_name)
        if not r:
            continue
        buy_cost = defaultdict(list)   # item_code -> [base_rate LYD]
        for p in pis:
            for it in p.get("items", []):
                if it.get("item_code") and it.get("base_rate"):
                    buy_cost[it["item_code"]].append(it["base_rate"])
        for it in r.get("items", []):
            code = it.get("item_code")
            sell = it.get("base_rate") or 0
            if code in buy_cost and sell > 0:
                avg_buy = statistics.mean(buy_cost[code])
                if sell < avg_buy:
                    flag_so(r, "سعر البيع أقل من سعر الشراء (خسارة)", "عالية",
                            f"صنف {code}: بيع {sell:,.2f} LYD < شراء (متوسط) {avg_buy:,.2f} LYD")

    # === فحص 9: طلبات مكررة (نفس العميل + نفس اليوم + نفس الإجمالي) ==========
    # ملاحظة: كل الطلبات تقريباً تستخدم نفس الصنف العام "سلة شي ان" (سلة/باقة
    # وليس منتجاً محدداً)، فالتكرار الحقيقي المشبوه هو نفس العميل + نفس اليوم
    # + نفس المبلغ الإجمالي بالضبط (احتمال إدخال الطلب مرتين بالخطأ).
    dup_groups = defaultdict(list)
    for r in so:
        key = (r.get("customer"), r.get("transaction_date"), round(r.get("grand_total") or 0, 2))
        dup_groups[key].append(r["name"])
    for (customer, d, total), names in dup_groups.items():
        uniq = sorted(set(names))
        if len(uniq) > 1 and total > 0:
            for n in uniq:
                r = so_by_name[n]
                flag_so(r, "طلب مكرر محتمل (نفس العميل + اليوم + الإجمالي)", "متوسطة",
                        f"العميل {customer}، التاريخ {d}، الإجمالي {total:,.2f} {r.get('currency')}، "
                        f"الطلبات: {', '.join(uniq)}")

    # === فحص 10: حقول ناقصة ================================================
    required = ["customer", "sales_partner", "transaction_date"]
    for r in so:
        missing = [f for f in required if not r.get(f)]
        if not r.get("items"):
            missing.append("items")
        if missing:
            flag_so(r, "حقول ناقصة في طلب البيع", "متوسطة", f"الحقول الناقصة: {', '.join(missing)}")

    # === فحص 11: القيم الشاذة (وسيط ± 3×MAD) لكل صنف + وكيل ==================
    groups_rate = defaultdict(list)
    groups_qty = defaultdict(list)
    item_rows = []  # (so_rec, item, key)
    for r in so:
        agent = r.get("sales_partner") or "(بدون وكيل)"
        for it in r.get("items", []):
            key = (it.get("item_code"), agent)
            groups_rate[key].append(it.get("rate"))
            groups_qty[key].append(it.get("qty"))
            item_rows.append((r, it, key))

    rate_bounds = {k: mad_bounds(v, ITEM_MAD_K) for k, v in groups_rate.items()}
    qty_bounds = {k: mad_bounds(v, ITEM_MAD_K) for k, v in groups_qty.items()}

    for r, it, key in item_rows:
        rb = rate_bounds.get(key)
        if rb and it.get("rate") is not None and not (rb[0] <= it["rate"] <= rb[1]):
            flag_so(r, "سعر شاذ لهذا الصنف/الوكيل (وسيط±3MAD)", "متوسطة",
                    f"صنف {it.get('item_code')}: rate={it['rate']} خارج [{rb[0]:.2f}, {rb[1]:.2f}] (الوسيط {rb[2]:.2f})")
        qb = qty_bounds.get(key)
        if qb and it.get("qty") is not None and not (qb[0] <= it["qty"] <= qb[1]):
            flag_so(r, "كمية شاذة لهذا الصنف/الوكيل (وسيط±3MAD)", "متوسطة",
                    f"صنف {it.get('item_code')}: qty={it['qty']} خارج [{qb[0]:.2f}, {qb[1]:.2f}] (الوسيط {qb[2]:.2f})")

    # === الحفظ ================================================================
    Path("data").mkdir(exist_ok=True)
    flags.save("data/flags.csv")

    agent_rows = []
    for agent, total in total_docs_by_agent.items():
        flagged = len(flagged_docs_by_agent.get(agent, ()))
        agent_rows.append({"agent": agent, "total_orders": total, "flagged_orders": flagged,
                            "pct_flagged": round(flagged / total * 100, 1) if total else 0})
    agent_rows.sort(key=lambda x: -x["pct_flagged"])
    with open("data/agent_summary.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["agent", "total_orders", "flagged_orders", "pct_flagged"])
        w.writeheader()
        w.writerows(agent_rows)

    print(f"إجمالي الحالات المشبوهة: {len(flags.rows)}")
    print(f"إجمالي الطلبات المعلَّمة (SO فريدة): {sum(len(v) for v in flagged_docs_by_agent.values())}")
    by_reason = Counter(x["reason"] for x in flags.rows)
    print("\nحسب نوع الفحص:")
    for reason, n in by_reason.most_common():
        print(f"  {n:>6}  {reason}")
    print(f"\nsaved -> data/flags.csv ({len(flags.rows)} rows), data/agent_summary.csv ({len(agent_rows)} agents)")


if __name__ == "__main__":
    main()
