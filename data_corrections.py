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
    # الشحنة JTE300464647589 (فاتورة الشراء ACC-PINV-2026-08907، جزء من مجموعة
    # الـ23 فاتورة المُعاد ربطها بـSAL-ORD-2026-05368): weight=125.0 كجم — على
    # الأرجح خطأ فاصلة عشرية ×100 (نفس نمط خطأ Mawadda). الدليل: شحنة شقيقة
    # بنفس رقم التتبع تقريباً JTE300464647589-1 (لفاتورة أخرى من نفس المجموعة
    # ACC-PINV-2026-08906) لها نفس cost=5.5 وexchange_rate=8.5 لكن weight=1.25.
    # القيمة المُصحَّحة = cost × 1.25، محوَّلة بنفس exchange_rate الأصلي (8.5).
    "JTE300464647589": {
        "weight": 1.25,
        "total_cost": 6.875,          # 5.5 × 1.25
        "total_amount": 6.875,
        "company_currency_total_cost": 58.4375,   # 6.875 × 8.5
        "company_currency_total_amount": 58.4375,
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


# SAL-ORD-2026-05236 كان طلباً خاطئاً (grand_total=1.00 LYD) للعميلة "الاء
# الغراري"، أُعيد إدخاله بشكل صحيح بعد 3 أيام كطلب جديد SAL-ORD-2026-05368
# (9,295 LYD) بدل تعديل الطلب الأصلي (استُبعد SAL-ORD-2026-05236 نفسه في
# exclusions.py). لكن هذه الـ23 فاتورة شراء الحقيقية (5,705.62 LYD) بقيت
# مربوطة في ERPNext بحقل sales_order إلى الطلب الخاطئ القديم — بينما سجلات
# الشحن المرتبطة بها تشير نصياً للطلب الصحيح 05368، ما يؤكد أنها فعلاً
# تخص الطلب الجديد. تصحيح محلي فقط (لا يلمس ERPNext) بطلب المستخدم.
PURCHASE_INVOICE_FIXES = {
    name: {"sales_order": "SAL-ORD-2026-05368"} for name in [
        "ACC-PINV-2026-08792", "ACC-PINV-2026-08793", "ACC-PINV-2026-08794",
        "ACC-PINV-2026-08795", "ACC-PINV-2026-08796", "ACC-PINV-2026-08798",
        "ACC-PINV-2026-08799", "ACC-PINV-2026-08800-1", "ACC-PINV-2026-08802",
        "ACC-PINV-2026-08803", "ACC-PINV-2026-08804", "ACC-PINV-2026-08807",
        "ACC-PINV-2026-08809", "ACC-PINV-2026-08810", "ACC-PINV-2026-08811",
        "ACC-PINV-2026-08812", "ACC-PINV-2026-08838", "ACC-PINV-2026-08840",
        "ACC-PINV-2026-08842", "ACC-PINV-2026-08844", "ACC-PINV-2026-08846",
        "ACC-PINV-2026-08906", "ACC-PINV-2026-08907",
    ]
}


def apply(so_records, sh_records=None, pi_records=None):
    """يُعدِّل so_records/sh_records/pi_records في مكانها (in place) ويُعيدها للتسلسل."""
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

    if pi_records is not None:
        for r in pi_records:
            fix = PURCHASE_INVOICE_FIXES.get(r.get("name"))
            if fix:
                r.update(fix)

    return so_records, sh_records, pi_records
