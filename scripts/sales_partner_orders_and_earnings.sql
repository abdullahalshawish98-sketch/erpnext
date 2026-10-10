-- طريقة التركيب في ERPNext:
--   Report List > New Report
--   Report Name         : Sales Partner Orders and Earnings  (أو أي اسم آخر غير مستخدم)
--   Report Type         : Query Report
--   Ref DocType         : Sales Order
--   Module              : Selling
--   Add Total Row        : ✅ فعّله (يضيف صف الإجمالي أسفل عمود "المكسب" تلقائياً)
--
--   الصق الاستعلام أدناه في حقل Query.
--
--   في جدول Filters أضف صفّين (Add Row مرتين):
--     1) Fieldname=sales_partner | Label=الوكيل      | Fieldtype=Link | Options=Sales Partner | Mandatory=✅
--     2) Fieldname=company       | Label=Company     | Fieldtype=Link | Options=Company       | Mandatory=✅
--   جعل "الوكيل" إلزامياً (Mandatory) هو ما يمنع ظهور أي نتيجة قبل اختيار
--   الوكيل أولاً، تماماً كما طلبت.
--
--   احفظ ثم افتح رابط التقرير — لن تظهر نتائج حتى تختار الوكيل والشركة.

SELECT
    so.name AS "رقم الطلب:Link/Sales Order:150",
    so.customer AS "العميل:Data:200",
    so.transaction_date AS "التاريخ:Date:110",
    so.grand_total AS "قيمة الطلب:Currency:130",
    so.total_commission AS "المكسب:Currency:130"
FROM `tabSales Order` so
WHERE so.docstatus = 1
    AND so.status = 'Completed'
    AND so.sales_partner = %(sales_partner)s
    AND so.company = %(company)s
ORDER BY so.transaction_date DESC
