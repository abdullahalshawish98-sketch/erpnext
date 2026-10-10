# طريقة التركيب في ERPNext:
#   Setup > Server Script > New
#   Script Type          : DocType Event
#   Reference Document Type : Purchase Invoice
#   DocType Event         : Before Submit
#   الصق الكود أدناه في حقل Script، ثم فعّل (Enable) واحفظ.
#
# الفحص: "سعر البيع أقل من سعر الشراء (خسارة)" — لكل صنف في فاتورة الشراء
# نقارن base_rate (تكلفة الشراء بعملة الشركة LYD) بسعر بيعه في طلب البيع
# المرتبط (Sales Order Item.base_rate لنفس item_code). تنبيه فقط، بدون حظر.

if doc.sales_order:
    so_items = frappe.get_all(
        "Sales Order Item",
        filters={"parent": doc.sales_order},
        fields=["item_code", "base_rate"],
    )
    sell_price = {}
    for it in so_items:
        if it.item_code:
            sell_price[it.item_code] = it.base_rate or 0

    messages = []
    for it in doc.items:
        sell = sell_price.get(it.item_code)
        buy = it.base_rate or 0
        if sell is not None and buy > 0 and sell < buy:
            messages.append(
                "صنف {0}: سعر البيع في {1} هو {2:,.2f} LYD < تكلفة الشراء {3:,.2f} LYD".format(
                    it.item_code, doc.sales_order, sell, buy
                )
            )

    if messages:
        frappe.msgprint(
            "تنبيه خسارة — سعر البيع أقل من سعر الشراء:<br>" + "<br>".join(messages),
            title="تنبيه: بيع أقل من التكلفة",
            indicator="red",
            alert=True,
        )
