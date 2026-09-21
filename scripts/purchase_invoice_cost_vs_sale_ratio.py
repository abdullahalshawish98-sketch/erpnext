# طريقة التركيب في ERPNext:
#   Setup > Server Script > New
#   Script Type          : DocType Event
#   Reference Document Type : Purchase Invoice
#   DocType Event         : Before Submit
#   الصق الكود أدناه في حقل Script، ثم فعّل (Enable) واحفظ.
#
# الفحص: نسبة (إجمالي فواتير الشراء المرتبطة بنفس طلب البيع / إجمالي طلب
# البيع) بعملة الشركة (LYD). ثلاث درجات:
#   - نسبة >= 100%                  -> أحمر: "خسارة مؤكدة"
#   - نسبة < 15%                    -> أحمر: "منخفضة جداً" (تحقق من اكتمال الفواتير)
#   - نسبة خارج [43%, 94%] (المدى الطبيعي وسيط±3MAD) لكن >= 15% -> برتقالي: "شاذة"
# تنبيه فقط (لا حظر تلقائي) — القرار للمستخدم دائماً.
#
# ⚠️ الحدود أدناه (LOW_RATIO / HIGH_RATIO) محسوبة من بيانات حتى 2026-09-20
# عبر reconcile.py. أعد حسابها دورياً (كل بضعة أشهر أو إذا تغيّر هامش الربح
# المعتاد للشركة) وحدّث القيم هنا يدوياً.

LOW_RATIO = 0.4317   # الوسيط 0.6865 ناقص 3×MAD
HIGH_RATIO = 0.9413  # الوسيط 0.6865 زائد 3×MAD
VERY_LOW_RATIO = 0.15  # تحت هذا الحد نعتبرها أخطر من مجرد "شاذة" (تنبيه أحمر)

if doc.sales_order:
    so = frappe.db.get_value(
        "Sales Order", doc.sales_order, ["name", "base_grand_total"], as_dict=True
    )
    if so and so.base_grand_total:
        prior_cost = frappe.db.sql(
            """
            select sum(base_grand_total)
            from `tabPurchase Invoice`
            where sales_order = %s and docstatus = 1 and name != %s
            """,
            (doc.sales_order, doc.name or ""),
        )[0][0] or 0
        total_cost = prior_cost + (doc.base_grand_total or 0)
        ratio = total_cost / so.base_grand_total

        if ratio >= 1:
            frappe.msgprint(
                "خسارة مؤكدة: إجمالي تكلفة الشراء المرتبطة بـ {0} هو {1:,.2f} "
                "وهو >= سعر بيع الطلب {2:,.2f} (نسبة التكلفة {3:.0f}%).".format(
                    so.name, total_cost, so.base_grand_total, ratio * 100
                ),
                title="تنبيه: خسارة مؤكدة",
                indicator="red",
                alert=True,
            )
        elif ratio < VERY_LOW_RATIO:
            frappe.msgprint(
                "نسبة تكلفة الشراء إلى سعر البيع منخفضة جداً لطلب البيع {0}: {1:.0f}% "
                "(أقل من {2:.0f}%) — تحقق من اكتمال فواتير الشراء المرتبطة أو صحة الربط بالطلب.".format(
                    so.name, ratio * 100, VERY_LOW_RATIO * 100
                ),
                title="تنبيه: نسبة تكلفة منخفضة جداً",
                indicator="red",
                alert=True,
            )
        elif ratio < LOW_RATIO or ratio > HIGH_RATIO:
            note = "مرتفعة (هامش ربح ضعيف)" if ratio > HIGH_RATIO else "منخفضة (تحقق من اكتمال فواتير الشراء)"
            frappe.msgprint(
                "نسبة تكلفة الشراء إلى سعر البيع {0} لطلب البيع {1}: {2:.0f}% "
                "(المدى الطبيعي المعتاد {3:.0f}%–{4:.0f}%).".format(
                    note, so.name, ratio * 100, LOW_RATIO * 100, HIGH_RATIO * 100
                ),
                title="تنبيه: نسبة تكلفة شاذة",
                indicator="orange",
                alert=True,
            )
