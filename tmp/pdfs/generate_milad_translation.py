from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph


OUT = Path("output/pdf/milad_hospital_translation.pdf")


PAGES = [
    {
        "page": "1",
        "patient": "Ms. Zohreh Khanom Eftekhari",
        "father": "Mirza Ali Akbar",
        "diagnosis": "sepsis and renal failure",
        "start": "1401/06/20",
        "end": "1401/06/28",
        "ward": "ICU",
        "doctor": "Dr. Saba",
    },
    {
        "page": "2",
        "patient": "Ms. Zohreh Khanom Eftekhari",
        "father": "Mirza Ali Akbar",
        "diagnosis": "sepsis",
        "start": "1401/01/21",
        "end": "1401/01/25",
        "ward": "ICU",
        "doctor": "Dr. Saba",
    },
    {
        "page": "3",
        "patient": "Ms. Zohreh Khanom Eftekhari",
        "father": "Mirza Ali Akbar",
        "diagnosis": "sepsis",
        "start": "1401/06/13",
        "end": "1401/06/19",
        "ward": "ICU",
        "doctor": "Dr. Saba",
    },
]


def centered(c, text, y, font="Times-Bold", size=18):
    width, _height = A4
    c.setFont(font, size)
    c.drawCentredString(width / 2, y, text)


def draw_wrapped(c, text, x, y, width, style):
    paragraph = Paragraph(text, style)
    _w, h = paragraph.wrap(width, 200 * mm)
    paragraph.drawOn(c, x, y - h)
    return y - h


def draw_stamp(c, x, y):
    c.saveState()
    c.setStrokeColor(colors.HexColor("#2D77C8"))
    c.setFillColor(colors.HexColor("#2D77C8"))
    c.setLineWidth(1.4)
    c.circle(x + 29 * mm, y + 22 * mm, 24 * mm)
    c.setFont("Helvetica-Bold", 9)
    c.drawCentredString(x + 29 * mm, y + 31 * mm, "STAMP")
    c.setFillColor(colors.black)
    c.setFont("Times-Bold", 11)
    c.drawString(x + 3 * mm, y + 12 * mm, "Health Information Management Department")
    c.drawString(x + 3 * mm, y + 3 * mm, "Milad Hospital Kashan")
    c.restoreState()


def draw_signature(c, x, y):
    c.setStrokeColor(colors.HexColor("#203A7A"))
    c.setLineWidth(1.0)
    c.line(x, y + 14 * mm, x + 65 * mm, y + 22 * mm)
    c.setFillColor(colors.black)
    c.setFont("Times-Italic", 10)
    c.drawString(x + 4 * mm, y + 4 * mm, "Signature")


def add_label_value(c, label, value, x, y, label_w=38 * mm):
    c.setFont("Times-Bold", 12)
    c.drawString(x, y, label)
    c.setFont("Times-Roman", 12)
    c.drawString(x + label_w, y, value)


def draw_page(c, data):
    width, height = A4
    margin = 18 * mm
    inner_x = margin
    inner_w = width - 2 * margin

    c.setStrokeColor(colors.black)
    c.setLineWidth(1)
    c.rect(margin, margin, inner_w, height - 2 * margin)

    c.setFont("Times-Bold", 11)
    c.drawString(margin + 6 * mm, height - margin - 12 * mm, "RM-FO-03")
    c.setFont("Helvetica", 10)
    c.drawRightString(width - margin - 6 * mm, height - margin - 12 * mm, "English Translation")

    centered(c, "Milad Hospital", height - margin - 34 * mm, size=22)
    centered(c, "In the Name of God", height - margin - 63 * mm, font="Times-Bold", size=13)

    style = ParagraphStyle(
        "body",
        fontName="Times-Roman",
        fontSize=13,
        leading=19,
        alignment=0,
        textColor=colors.black,
    )

    body = (
        f"Respectfully, this is to certify that {data['patient']}, daughter of "
        f"{data['father']}, with the diagnosis of {data['diagnosis']}, was hospitalized "
        f"from {data['start']} to {data['end']} in the {data['ward']} ward of Milad "
        f"Hospital. The treating physician was {data['doctor']}. This certificate has "
        "been issued at the request of the above-named person for presentation and has "
        "no other validity."
    )
    box_x = inner_x + 20 * mm
    y = height - margin - 105 * mm
    draw_wrapped(c, body, box_x, y, inner_w - 40 * mm, style)

    draw_signature(c, box_x + 70 * mm, margin + 48 * mm)
    draw_stamp(c, box_x, margin + 28 * mm)

    c.setFont("Times-Roman", 9)
    footer = "Stamp text translated according to the visible seal: Health Information Management Department / Milad Hospital Kashan"
    footer_w = stringWidth(footer, "Times-Roman", 9)
    c.drawString((width - footer_w) / 2, margin + 8 * mm, footer)
    c.showPage()


def main():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(OUT), pagesize=A4)
    c.setTitle("Milad Hospital English Translation")
    for page in PAGES:
        draw_page(c, page)
    c.save()
    print(OUT)


if __name__ == "__main__":
    main()
