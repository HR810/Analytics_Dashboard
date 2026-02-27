import smtplib
from email.message import EmailMessage
from reportlab.platypus import PageBreak
from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer,
        Image, Table, TableStyle
    )
from datetime import datetime
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
from reportlab.lib.units import inch

def generate_pdf_report(
    fig_bar,
    fig_month_client,
    fig_line,
    fig_heatmap,
    kpis,
    clients,
    assignees,
    months,
    weeks,
    filtered_df
):

    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer,
        Image, Table, TableStyle, PageBreak
    )
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import inch
    from datetime import datetime
    from io import BytesIO

    # =========================
    # Footer
    # =========================
    def add_footer(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 9)

        timestamp = f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M')}"

        canvas.drawString(40, 20, timestamp)
        canvas.drawRightString(A4[0] - 40, 20, f"Page {doc.page}")

        canvas.restoreState()

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    elements = []
    styles = getSampleStyleSheet()

    section_style = ParagraphStyle(
        name='SectionTitle',
        parent=styles['Heading2'],
        fontSize=16,
        spaceAfter=10
    )

    wrap_style = ParagraphStyle(
        name='WrapStyle',
        parent=styles['Normal'],
        fontSize=10,
        leading=12
    )

    # =========================
    # TITLE
    # =========================
    elements.append(Paragraph("Customer Ticket Effort Report", styles["Title"]))
    elements.append(Spacer(1, 0.3 * inch))

    # =========================
    # SUMMARY METRICS
    # =========================
    elements.append(Paragraph("Summary Metrics", section_style))

    summary_data = [
        ["Metric", "Value"],
        ["Total Effort", kpis["effort"]],
        ["Total Tickets", kpis["tickets"]],
        ["Active Assignees", kpis["assignees"]],
    ]

    summary_table = Table(summary_data, colWidths=[3 * inch, 2 * inch])
    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('ALIGN', (1, 1), (-1, -1), 'RIGHT')
    ]))

    elements.append(summary_table)
    elements.append(Spacer(1, 0.4 * inch))

    # =========================
    # FILTERS
    # =========================
    elements.append(Paragraph("Applied Filters", section_style))

    filter_data = [
        ["Client", Paragraph(", ".join(clients), wrap_style)],
        ["Assignee", Paragraph(", ".join(assignees), wrap_style)],
        ["Month", Paragraph(", ".join(months), wrap_style)],
        ["Week", Paragraph(", ".join(weeks), wrap_style)],
    ]

    filter_table = Table(filter_data, colWidths=[1.5 * inch, 4.5 * inch])
    filter_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, -1), colors.whitesmoke),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]))

    elements.append(filter_table)
    elements.append(PageBreak())

    # =========================
    # CLIENT SUMMARY TABLE
    # =========================
    elements.append(Paragraph("Total Effort per Client", section_style))

    client_summary = (
        filtered_df.groupby("Client")["Effort"]
        .sum()
        .sort_values(ascending=False)
        .reset_index()
    )

    client_data = [["Client", "Total Effort"]]
    for _, row in client_summary.iterrows():
        client_data.append([row["Client"], int(row["Effort"])])

    client_table = Table(client_data, colWidths=[3 * inch, 2 * inch])
    client_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('ALIGN', (1, 1), (-1, -1), 'RIGHT')
    ]))

    elements.append(client_table)
    elements.append(Spacer(1, 0.4 * inch))

    # =========================
    # ASSIGNEE SUMMARY TABLE
    # =========================
    elements.append(Paragraph("Total Effort per Assignee", section_style))

    assignee_summary = (
        filtered_df.groupby("Assignee")["Effort"]
        .sum()
        .sort_values(ascending=False)
        .reset_index()
    )

    assignee_data = [["Assignee", "Total Effort"]]
    for _, row in assignee_summary.iterrows():
        assignee_data.append([row["Assignee"], int(row["Effort"])])

    assignee_table = Table(assignee_data, colWidths=[3 * inch, 2 * inch])
    assignee_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('ALIGN', (1, 1), (-1, -1), 'RIGHT')
    ]))

    elements.append(assignee_table)
    elements.append(PageBreak())

    # =========================
    # EXPORT CHART IMAGES
    # =========================
    bar_png = fig_bar.to_image(format="png", scale=3)
    month_client_png = fig_month_client.to_image(format="png", scale=3)
    line_png = fig_line.to_image(format="png", scale=3)
    heat_png = fig_heatmap.to_image(format="png", scale=3)

    # =========================
    # ASSIGNEE BAR CHART
    # =========================
    elements.append(Paragraph("Monthly Total Effort per Assignee", section_style))
    elements.append(Image(BytesIO(bar_png), width=6 * inch, height=3 * inch))
    elements.append(Spacer(1, 0.4 * inch))

    # =========================
    # MONTHLY CLIENT CHART
    # =========================
    elements.append(Paragraph("Monthly Total Effort per Client", section_style))
    elements.append(Image(BytesIO(month_client_png), width=6 * inch, height=3 * inch))
    elements.append(PageBreak())

    # =========================
    # WEEKLY TREND
    # =========================
    elements.append(Paragraph("Weekly Effort Trend per Client", section_style))
    elements.append(Image(BytesIO(line_png), width=6 * inch, height=3 * inch))
    elements.append(Spacer(1, 0.4 * inch))

    # =========================
    # HEATMAP
    # =========================
    elements.append(Paragraph("Weekly Effort Heatmap per Client", section_style))
    elements.append(Image(BytesIO(heat_png), width=6 * inch, height=3 * inch))

    # =========================
    # BUILD PDF
    # =========================
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
