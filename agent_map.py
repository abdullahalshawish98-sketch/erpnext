"""
خريطة حرف بادئة اسم بطاقة الهدية ↔ الوكيل/المتجر الذي خُصِّصت له.
المستخدم أضاف هذه الحروف يدوياً (في اسم البطاقة الخارجي في ERPNext/شي ان)
عند توزيع كل بطاقة على وكيل. أكّد المستخدم الأسماء الحقيقية شفهياً؛
الحساب المهيمن مُستخرَج تلقائياً من حقل owner في فواتير الشراء المرتبطة
بكل بطاقة (نسبة الهيمنة بين قوسين).
"""

CARD_LETTER_TO_AGENT = {
    "S": "متجر سلتك",       # saltk@aqsaa.com (87%) + saltic@gmail.com (9%)
    "Z": "سيزون ستور",      # seasonstore@aqsaa.com (99%)
    "M": "متجر وصلها",      # waselha@aqsaa.com (98%)
    "Mm": "متجر وصلها",     # waselha@aqsaa.com (100%) — نفس المتجر
    "Y": "يوسف",            # yousefbuhjar7@gmail.com (99%)
    "N": "عدنان",            # lamoshaen@gmail.com (97%)
    "L": "عالم نون",         # ali1alshukri@gmail.com (94%)
    "H": "Libyol",           # moha@alaqsa.com (92%)
    "R": "لبابك",            # lebabek@aqsaa.com (98%)
    "G": "Blue store",       # naded@gezwi@gmail.com (99%)
    "I": "Lumar store",      # shipa@alaqsa.com (86%)
    "A": "عبدالله الشاوش (المسؤول)",  # abdo.z@gmail.com (68%) — طلبيات المسؤول نفسه، ليست وكيلاً خارجياً
}

# ليست وكلاء خارجيين — حسابات مكتب/إدارة داخلية
OFFICE_ACCOUNTS_BY_LETTER = {
    "Sm": "sadeg@aqsaa.ly — حساب مكتب/إدارة",
    "B": "zurgani@aqsaa.com — غالباً فريق شراء بالمكتب",
}
