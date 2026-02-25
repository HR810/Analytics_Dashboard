import os
import smtplib
from email.message import EmailMessage
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import Table
from reportlab.platypus import TableStyle
from io import BytesIO
import plotly.io as pio
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.units import inch
from io import BytesIO
from reportlab.platypus import PageBreak
from reportlab.lib.utils import ImageReader

def generate_pdf_report(fig_bar, fig_line, fig_heatmap,
                        kpis, clients, assignees, months, weeks):

    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer,
        Image, Table, TableStyle
    )
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    from reportlab.lib.units import inch
    from io import BytesIO
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import inch
    from datetime import datetime

    def add_footer(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 9)

        timestamp = f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M')}"

        # Bottom-left corner
        canvas.drawString(
            40,  # left margin
            20,  # distance from bottom
            timestamp
        )

        # Optional: Page number bottom-right
        page_number_text = f"Page {doc.page}"
        canvas.drawRightString(
            A4[0] - 40,
            20,
            page_number_text
        )

        canvas.restoreState()
    from datetime import datetime

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer)
    elements = []
    styles = getSampleStyleSheet()

    section_style = ParagraphStyle(
        name='SectionTitle',
        parent=styles['Heading2'],
        fontSize=16,
        spaceAfter=10
    )

    # =========================
    # REPORT TITLE
    # =========================
    elements.append(Paragraph("Customer Ticket Effort Report", styles["Title"]))
    elements.append(Spacer(1, 0.2 * inch))

    # =========================
    # SUMMARY TABLE
    # =========================
    elements.append(Paragraph("Summary Metrics", section_style))

    summary_data = [
        ["Metric", "Value"],
        ["Total Effort", kpis["effort"]],
        ["Total Tickets", kpis["tickets"]],
        ["Active Assignees", kpis["assignees"]],
    ]

    summary_table = Table(summary_data, colWidths=[200, 100])
    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica')
    ]))

    elements.append(summary_table)
    elements.append(Spacer(1, 0.4 * inch))

    # =========================
    # FILTER TABLE
    # =========================
    from reportlab.platypus import Paragraph
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib import colors
    from reportlab.platypus import Table, TableStyle
    from reportlab.lib.units import inch

    # Custom wrapping style
    wrap_style = ParagraphStyle(
        name='WrapStyle',
        parent=styles['Normal'],
        fontSize=10,
        leading=12
    )

    elements.append(Paragraph("Applied Filters", section_style))

    filter_data = [
        ["Client", Paragraph(", ".join(clients), wrap_style)],
        ["Assignee", Paragraph(", ".join(assignees), wrap_style)],
        ["Month", Paragraph(", ".join(months), wrap_style)],
        ["Week", Paragraph(", ".join(weeks), wrap_style)],
    ]

    filter_table = Table(filter_data, colWidths=[1.2 * inch, 4.8 * inch])

    filter_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, -1), colors.whitesmoke),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))

    elements.append(filter_table)
    elements.append(Spacer(1, 0.4 * inch))

    # =========================
    # EXPORT IMAGES
    # =========================
    bar_png = fig_bar.to_image(format="png", scale=3)
    line_png = fig_line.to_image(format="png", scale=3)
    heat_png = fig_heatmap.to_image(format="png", scale=3)

    # =========================
    # BAR CHART
    # =========================
    elements.append(Paragraph("Monthly Total Effort per Assignee", section_style))
    elements.append(Image(BytesIO(bar_png), width=6*inch, height=3*inch))
    elements.append(Spacer(1, 0.4 * inch))

    # =========================
    # LINE CHART
    # =========================
    elements.append(PageBreak())
    elements.append(Paragraph("Weekly Effort Trend per Client", section_style))
    elements.append(Image(BytesIO(line_png), width=6*inch, height=3*inch))
    elements.append(Spacer(1, 0.4 * inch))

    # =========================
    # HEATMAP
    # =========================
    elements.append(Paragraph("Weekly Effort Heatmap per Client", section_style))
    elements.append(Image(BytesIO(heat_png), width=6*inch, height=3*inch))

    doc.build(
        elements,
        onFirstPage=add_footer,
        onLaterPages=add_footer
    )
    buffer.seek(0)

    return buffer

def send_email_report(receiver_email, pdf_buffer):
    sender_email = "harisankars0810@gmail.com"
    sender_password = "qdgg yrml muoc slaw"  # Use Gmail App Password

    msg = EmailMessage()
    msg["Subject"] = "Customer Ticket Effort Report"
    msg["From"] = sender_email
    msg["To"] = receiver_email
    msg.set_content("Please find attached the filtered dashboard report.")

    msg.add_attachment(
        pdf_buffer.read(),
        maintype="application",
        subtype="pdf",
        filename="Ticket_Report.pdf"
    )

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(sender_email, sender_password)
        server.send_message(msg)
