import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch

# اسم العلامة المائية
watermark_text = "𝔼𝕃 ℍ𝕒𝕣𝕒𝕞𓂀"

# دالة لإضافة العلامة المائية
def add_watermark(canvas, doc):
    canvas.saveState()
    canvas.setFont('Helvetica-Bold', 40)
    canvas.setFillColorRGB(0.8, 0.8, 0.8, alpha=0.2)  # لون فاتح وشفاف
    canvas.translate(A4[0]/2, A4[1]/2)
    canvas.rotate(30)  # ميل بسيط
    canvas.drawCentredString(0, 0, watermark_text)
    canvas.restoreState()

# اقرأ بيانات من ملف Excel
df = pd.read_excel("students.xlsx")

# إنشاء ملف PDF
pdf_file = "students_report.pdf"
doc = SimpleDocTemplate(pdf_file, pagesize=A4)

# ستايلات جاهزة
styles = getSampleStyleSheet()
title_style = styles['Heading1']

# عنوان التقرير
title = Paragraph("📊 Students Report", title_style)

# تحويل الـ DataFrame لجدول
data = [df.columns.tolist()] + df.values.tolist()

table = Table(data)
table.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#4CAF50")),  # خلفية للهيدر
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('FONTSIZE', (0, 0), (-1, 0), 12),
    ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
    ('BACKGROUND', (0, 1), (-1, -1), colors.whitesmoke),
    ('GRID', (0, 0), (-1, -1), 1, colors.black),
]))

# تجميع العناصر
elements = [title, Spacer(1, 20), table]

# بناء الملف مع العلامة المائية
doc.build(elements, onFirstPage=add_watermark, onLaterPages=add_watermark)

print("✅ تقرير احترافي مع علامة مائية:", pdf_file)