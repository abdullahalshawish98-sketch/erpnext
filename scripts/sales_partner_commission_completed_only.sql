-- طريقة التركيب في ERPNext (لا يحتاج Developer Mode ولا وصولاً للملفات):
--   Report List > New Report
--   Report Name        : Sales Partner Commission Summary (Completed Only)
--   Report Type        : Query Report
--   Ref DocType        : Sales Order
--   Module             : Selling
--   الصق الاستعلام أدناه في حقل Query، ثم احفظ.
--
--   بعدها أضف فلتر Company (اختياري لكن يطابق التقرير الأصلي):
--   في تبويب Filters بمستند التقرير، أضف صفاً:
--     Fieldname = company   |   Label = Company   |   Fieldtype = Link   |   Options = Company
--
-- الفرق عن التقرير القياسي "Sales Partner Commission Summary": هذا يحسب
-- العمولة فقط لطلبات البيع المكتملة (status = 'Completed')، بدل كل الطلبات
-- المعتمدة بغض النظر عن حالتها — لم يُعدَّل التقرير الأصلي إطلاقاً، هذا تقرير
-- منفصل جديد.
--
-- ملاحظة: Frappe يمرّر هذا الاستعلام عبر تنسيق %-style من أجل حقن الفلاتر
-- (%(company)s)، لذا أي علامة % حرفية في اسم عمود يجب كتابتها %% (وليس %
-- مفردة) وإلا فشل الاستعلام بخطأ "unsupported format character".

SELECT
    so.sales_partner AS "الوكيل:Link/Sales Partner:200",
    COUNT(so.name) AS "عدد الطلبات المكتملة:Int:120",
    SUM(so.amount_eligible_for_commission) AS "المبلغ المؤهَّل للعمولة:Currency:180",
    AVG(so.commission_rate) AS "نسبة العمولة%%:Float:100",
    SUM(so.total_commission) AS "إجمالي العمولة:Currency:180"
FROM `tabSales Order` so
WHERE so.docstatus = 1
    AND so.status = 'Completed'
    AND so.company = %(company)s
GROUP BY so.sales_partner
ORDER BY SUM(so.total_commission) DESC
