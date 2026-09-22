"""
تصحيحات مؤقتة لأخطاء بيانات مؤكدة في ERPNext، بانتظار تصحيحها يدوياً هناك.
تُطبَّق على البيانات المحلية في data/*.json فقط (لا تلمس ERPNext نفسه).
احذف كل إدخال من هنا فور تصحيح المستند الأصلي في ERPNext حتى لا تُخفي
تصحيحاً حقيقياً لاحقاً بالخطأ.

قاعدة كل تصحيح: نص السبب + المصدر (كيف اكتُشف) حتى يسهل التحقق لاحقاً.
"""

# الشحنة 1-JTE300477453789 (SAL-ORD-2026-07103، الوكيل: متجر لبابك، العميل: Mawadda):
# total_cost محسوب خطأً بمقدار ×1000 (cost × weight = 5.5 × 1.4 = 7.70، والمُدخَل
# فعلياً 7700.00). القيمة المُصحَّحة أدناه = cost × weight بالضبط، محوَّلة بنفس
# exchange_rate الأصلي (8.75) لعمود company_currency_*.
SHIPMENT_FIXES = {
    "1-JTE300477453789": {
        "total_cost": 7.7,
        "total_amount": 7.7,
        "company_currency_total_cost": 67.375,   # 7.7 × 8.75
        "company_currency_total_amount": 67.375,
    },
}

# نفس الرقم الخاطئ (67375) انتقل يدوياً إلى بند "شحن" في طلب البيع نفسه
# (كان مُدخَلاً 67425 LYD بدل ~67.4 LYD) مما رفع grand_total الطلب زوراً من
# ~572 إلى 67930 LYD.
SALES_ORDER_ITEM_FIXES = {
    # اسم الطلب: {item_code: {rate/amount/base_rate/base_amount الصحيحة}}
    "SAL-ORD-2026-07103": {
        "شحن": {"rate": 67.375, "amount": 67.375, "base_rate": 67.375, "base_amount": 67.375},
    },
}


def apply(so_records, sh_records=None):
    """يُعدِّل so_records/sh_records في مكانها (in place) ويُعيدها للتسلسل."""
    item_fixes = SALES_ORDER_ITEM_FIXES
    for r in so_records:
        fix = item_fixes.get(r.get("name"))
        if not fix:
            continue
        for it in r.get("items", []):
            item_fix = fix.get(it.get("item_code"))
            if item_fix:
                it.update(item_fix)
        # أعد حساب إجمالي الطلب من بنوده بعد التصحيح بدل تخمين رقم جاهز
        r["grand_total"] = sum(it.get("amount") or 0 for it in r.get("items", []))
        r["base_grand_total"] = sum(it.get("base_amount") or 0 for it in r.get("items", []))

    if sh_records is not None:
        for r in sh_records:
            fix = SHIPMENT_FIXES.get(r.get("name"))
            if fix:
                r.update(fix)

    return so_records, sh_records
