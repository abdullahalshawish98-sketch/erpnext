# طريقة التركيب في ERPNext:
#   Setup > Server Script > New
#   Script Type          : DocType Event
#   Reference Document Type : Purchase Invoice
#   DocType Event         : Before Submit
#   الصق الكود أدناه في حقل Script، ثم فعّل (Enable) واحفظ.
#
# الفحص: "سعر صرف شاذ في فاتورة الشراء (وسيط±3MAD)". فواتير الشراء بعملة
# الشركة (LYD) لا تحتاج تحويلاً (conversion_rate = 1 طبيعي لها)، لذا نطبّق
# الفحص فقط على الفواتير بعملة مختلفة (USD في حالتنا). تنبيه فقط، بدون حظر.
#
# ⚠️ الحدود أدناه محسوبة من بيانات حتى 2026-09-20 عبر reconcile.py.
# أعد حسابها دورياً وحدّثها هنا يدوياً، خصوصاً إذا تغيّر سعر الصرف الرسمي
# بشكل دائم (وليس خطأ إدخال).

FOREIGN_CURRENCY = "USD"
LOW_RATE = 7.1078    # الوسيط 8.58 ناقص 3×MAD
HIGH_RATE = 10.0522  # الوسيط 8.58 زائد 3×MAD

if doc.currency == FOREIGN_CURRENCY and doc.conversion_rate:
    if doc.conversion_rate == 1:
        frappe.msgprint(
            "سعر الصرف = 1 رغم أن عملة الفاتورة {0} — يبدو أنه لم يُحدَّث عند الإدخال.".format(
                FOREIGN_CURRENCY
            ),
            title="تنبيه: سعر صرف لم يُحدَّث",
            indicator="orange",
            alert=True,
        )
    elif doc.conversion_rate < LOW_RATE or doc.conversion_rate > HIGH_RATE:
        frappe.msgprint(
            "سعر الصرف المُدخَل {0:.4f} خارج المدى الطبيعي المعتاد ({1:.2f}–{2:.2f}) "
            "لعملة {3} — تحقق من الرقم قبل الاعتماد.".format(
                doc.conversion_rate, LOW_RATE, HIGH_RATE, FOREIGN_CURRENCY
            ),
            title="تنبيه: سعر صرف شاذ",
            indicator="red",
            alert=True,
        )
