import requests
import os
import json
import logging
from io import BytesIO
import win32com.client as win32

from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.colors import HexColor
from reportlab.pdfgen import canvas
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageBreak, Flowable


# === Configuration ===
API_URL = "https://api.monday.com/v2"
API_TOKEN = "eyJhbGciOiJIUzI1NiJ9.eyJ0aWQiOjUyNTkwODExMiwiYWFpIjoxMSwidWlkIjo2ODEyMTYxMywiaWFkIjoiMjAyNS0wNi0xM1QwNDowMDoxNy4wMDBaIiwicGVyIjoibWU6d3JpdGUiLCJhY3RpZCI6MjYzMTYxNTksInJnbiI6ImFwc2UyIn0.l_VCsvsabCwF1gC55poSRvWJL9edZek6wXtyeg8kMuY"
BOARD_ID = 1955440067 
PREVIOUS_DATA_FILE = "ECO_previous_IDs.json"
PDF_OUTPUT_FILE = "ECO_Status_Change.pdf"
EMAIL_TO = "owen@regentrv.com.au"

headers = {
    "Authorization": API_TOKEN,
    "Content-Type": "application/json"
}

# === Logging Setup ===
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler("ECO_pdf_email.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("ECOBoardNameMonitor")

def draw_watermark(canvas, doc):
    try:
        # Load local image
        img_path = "logo.png"  # Make sure it's in the same folder
        img = ImageReader(img_path)

        # Image size (adjust as needed)
        img_width = 1.2 * inch
        img_height = 1.2 * inch

        # Top-right corner, inside the margins
        x = doc.pagesize[0] - doc.rightMargin - img_width
        y = doc.pagesize[1] - doc.topMargin - 30  # Push slightly above top margin if needed

        # Draw image
        canvas.drawImage(
            img, x, y,
            width=img_width,
            height=img_height,
            preserveAspectRatio=True,
            mask='auto'
        )

    except Exception as e:
        print("Failed to draw watermark:", e)


def load_previous_ids(filename):
    try:
        with open(filename, "r", encoding="utf-8") as f:
            return set(json.load(f))
    except (FileNotFoundError, json.JSONDecodeError):
        return set()

# === Fetch items from Monday.com ===
import json

def get_board_items(previous_ids):
    query = f"""
    {{
      boards(ids: [{BOARD_ID}]) {{
        items_page(limit: 10) {{
          items {{
            id
            name
            column_values {{
              column {{ title }}
              text
              value
            }}
            assets {{
              id
              name
              public_url
            }}
          }}
        }}
      }}
    }}
    """
    response = requests.post(API_URL, json={"query": query}, headers=headers)
    response.raise_for_status()
    data = response.json()
    items = data["data"]["boards"][0]["items_page"]["items"]

    new_started_items = []

    for item in items:

        # if item["id"] in previous_ids:
        #     continue

        status_col = next((col for col in item["column_values"] if col["column"]["title"] == "Status"), None)
        if not status_col or status_col["text"] != "Started":
            continue

        # Build a mapping of assetId to public_url
        asset_map = {int(asset["id"]): asset["public_url"] for asset in item.get("assets", [])}

        def extract_public_urls(column_title):
            col = next((c for c in item["column_values"] if c["column"]["title"] == column_title), None)
            if not col or not col.get("value"):
                return []
            try:
                value_json = json.loads(col["value"])
                return [
                    asset_map.get(file["assetId"])
                    for file in value_json.get("files", [])
                    if file["assetId"] in asset_map
                ]
            except (json.JSONDecodeError, KeyError, TypeError):
                return []

        # Extract Old and New Picture URLs
        old_picture_urls = extract_public_urls("ECR Old Picture")
        new_picture_urls = extract_public_urls("ECR New Picture")

        item["old_picture_urls"] = old_picture_urls
        item["new_picture_urls"] = new_picture_urls

        new_started_items.append(item)

    return new_started_items

# === Generate PDF ===
def generate_pdf(data, filename):
    

    # Set fixed margins
    margin = 0.75 * inch  # 0.75 inch on all sides (~54 points)
    available_width = A4[0] - 2 * margin

    doc = SimpleDocTemplate(
        filename,
        pagesize=A4,
        leftMargin=margin,
        rightMargin=margin,
        topMargin=margin,
        bottomMargin=margin
    )

    story = []
    styles = getSampleStyleSheet()
    styleN = styles["BodyText"]
    styleB = styles["Heading2"]

    if not data:
        story.append(Paragraph("No changes detected.", styleN))
    else:
        for item in data:

            # === Title ===
            title_style = styles["Title"]
            title_style.fontSize = 18
            title_style.alignment = 0  # Left aligned to control table layout properly

            # Title paragraph
            title_paragraph = Paragraph(f"{item.get('name', '')}", title_style)

            # ECO tag inside a green box
            eco_text = Paragraph(
                '<font color="#008000"><b>ECO</b></font>', styles["BodyText"]
            )
            eco_box = Table([[eco_text]], colWidths=30, rowHeights=20)
            eco_box.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, -1), colors.white),
                ('TEXTCOLOR', (0, 0), (-1, -1), colors.HexColor('#008000')),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#008000')),
                ('LEFTPADDING', (0, 0), (-1, -1), 4),
                ('RIGHTPADDING', (0, 0), (-1, -1), 4),
                ('TOPPADDING', (0, 0), (-1, -1), 0),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
            ]))

            # Combine title and ECO into a single row (no gap between)
            title_eco_table = Table(
                [[title_paragraph, eco_box]],
                colWidths=[available_width - 40, 40],  # Adjust width if necessary
            )
            title_eco_table.setStyle(TableStyle([
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 0),
                ('RIGHTPADDING', (0, 0), (-1, -1), 0),
                ('TOPPADDING', (0, 0), (-1, -1), 0),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
            ]))

            # Add to story
            story.append(title_eco_table)
            story.append(Spacer(1, 20))



            # === Section 0: Priority, no, date and category ===
            priority = next((col["text"] for col in item.get("column_values", []) 
                             if col["column"]["title"] == "Priority" and col.get("text")), "N/A")
            change_no = next((col["text"] for col in item.get("column_values", []) 
                              if col["column"]["title"] == "NO." and col.get("text")), "N/A")
            eco_date = next((col["text"] for col in item.get("column_values", []) 
                             if col["column"]["title"] == "ECO Date" and col.get("text")), "N/A")
            category = next((col["text"] for col in item.get("column_values", []) 
                             if col["column"]["title"] == "Category" and col.get("text")), "N/A")

            # Color for Priority
            priority_raw = next((col["text"] for col in item.get("column_values", []) 
                                if col["column"]["title"] == "Priority" and col.get("text")), "N/A").lower()

            if "critical" in priority_raw:
                priority_color = colors.red
                priority_text = "Critical"
            elif "high" in priority_raw:
                priority_color = colors.HexColor("#003366")  # Dark Blue
                priority_text = "High"
            elif "medium" in priority_raw:
                priority_color = colors.blue
                priority_text = "Medium"
            elif "low" in priority_raw:
                priority_color = colors.HexColor("#99CCFF")  # Light Blue
                priority_text = "Low"
            else:
                priority_color = colors.grey
                priority_text = priority_raw or "N/A"

            priority_para = Paragraph(
                f'<font color="white"><b>{priority_text}</b></font>', 
                ParagraphStyle('priority_style', alignment=1, fontSize=9)
            )

            priority_table = Table([[priority_para]], colWidths=[60], rowHeights=[20])
            priority_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, -1), priority_color),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('BOX', (0, 0), (-1, -1), 1, priority_color),
                ('TEXTCOLOR', (0, 0), (-1, -1), colors.white),
            ]))

            no_text = next((col["text"] for col in item.get("column_values", []) 
                            if col["column"]["title"] == "NO." and col.get("text") and col.get("text")), "N/A")

            no_para = Paragraph(f"NO. {no_text}", ParagraphStyle('no_style', fontSize=9, alignment=1, textColor=colors.grey))
            no_table = Table([[no_para]], colWidths=[90], rowHeights=[20])
            no_table.setStyle(TableStyle([
                ('BOX', (0, 0), (-1, -1), 1, colors.grey),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('TEXTCOLOR', (0, 0), (-1, -1), colors.grey),
            ]))


            eco_date_text = next((col["text"] for col in item.get("column_values", []) 
                                  if col["column"]["title"] == "ECO Date" and col.get("text")), "N/A")
            eco_date_para = Paragraph(eco_date_text, ParagraphStyle('eco_style', fontSize=9, alignment=1, textColor=colors.grey))
            eco_date_table = Table([[eco_date_para]], colWidths=[110], rowHeights=[20])
            eco_date_table.setStyle(TableStyle([
                ('BOX', (0, 0), (-1, -1), 1, colors.grey),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ]))


            category_text = next((col["text"] for col in item.get("column_values", []) 
                                  if col["column"]["title"] == "Category" and col.get("text")), "N/A")
            category_para = Paragraph(category_text, ParagraphStyle('cat_style', fontSize=9, alignment=1, textColor=colors.grey))
            category_table = Table([[category_para]], colWidths=[available_width - (60+90+110)], rowHeights=[20])
            category_table.setStyle(TableStyle([
                ('BOX', (0, 0), (-1, -1), 1, colors.grey),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ]))


            row_table = Table(
                [[priority_table, no_table, eco_date_table, category_table]],
                colWidths=[60, 90, 110, available_width - (60 + 90 + 110)],
                style=TableStyle([
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    ('LEFTPADDING', (0, 0), (-1, -1), 0),
                ])
            )

            story.append(row_table)
            story.append(Spacer(1, 10))





            # === Section 1: Description and Reason ==
            desc_text = next((col["text"] for col in item.get("column_values", []) 
                              if col["column"]["title"] == "ECR Description" and col.get("text")), "N/A")


            reason_text = next((col["text"] for col in item.get("column_values", []) 
                             if col["column"]["title"] == "ECR Reason" and col.get("text")),"N/A")

            # Custom paragraph styles
            desc_style = ParagraphStyle('desc_style', parent=styleN, fontSize=10, leading=14)
            bold_style = ParagraphStyle('bold_style', parent=styleN, fontSize=10, leading=14, spaceAfter=6, fontName='Helvetica-Bold')

            desc_para = Paragraph(f"<b>Change Description</b><br/>{desc_text}", desc_style)
            reason_para = Paragraph(f"<b>Change Reason</b><br/>{reason_text}", desc_style)

            # Wrap in a table with border and padding to simulate rounded box
            desc_table = Table([[desc_para], [reason_para]],
                colWidths=[available_width],
                style=TableStyle([
                    ('BOX', (0, 0), (-1, -1), 1, colors.grey),
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    ('BACKGROUND', (0, 0), (-1, -1), colors.white),
                    ('TOPPADDING', (0, 0), (-1, -1), 8),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
                ])
            )
            story.append(desc_table)
            story.append(Spacer(1, 20))




            # === Section 2: Images ===
            # Load old images into a vertical list of Image or Paragraph (if failed)
            old_images = [Paragraph('<b>Old Picture</b>', styleN), Spacer(1, 6)]
            for url in item.get("old_picture_urls", []):
                try:
                    response = requests.get(url)
                    if response.status_code == 200:
                        img_data = BytesIO(response.content)
                        img = Image(img_data, width=available_width/2 - 20, height=120)
                        old_images.append(img)
                        old_images.append(Spacer(1,6))
                except:
                    old_images.append(Paragraph("Image load failed", styleN))

            if len(old_images) == 2:  # Only title and spacer present, no images loaded
                old_images.append(Paragraph("N/A", styleN))

            old_box = Table([[old_images]], colWidths=[available_width/2])
            old_box.setStyle(TableStyle([
                ('BOX', (0,0), (-1,-1), 1, colors.grey),
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('TOPPADDING', (0,0), (-1,-1), 6),
                ('BOTTOMPADDING', (0,0), (-1,-1), 6),
            ]))

            # Prepare new images content with title inside the box
            new_images = [Paragraph('<b>New Picture</b>', styleN), Spacer(1, 6)]
            for url in item.get("new_picture_urls", []):
                try:
                    response = requests.get(url)
                    if response.status_code == 200:
                        img_data = BytesIO(response.content)
                        img = Image(img_data, width=available_width/2 - 20, height=120)
                        new_images.append(img)
                        new_images.append(Spacer(1,6))
                except:
                    new_images.append(Paragraph("Image load failed", styleN))

            if len(new_images) == 2:
                new_images.append(Paragraph("N/A", styleN))

            new_box = Table([[new_images]], colWidths=[available_width/2])
            new_box.setStyle(TableStyle([
                ('BOX', (0,0), (-1,-1), 1, colors.grey),
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('TOPPADDING', (0,0), (-1,-1), 6),
                ('BOTTOMPADDING', (0,0), (-1,-1), 6),
            ]))

            # Final table side by side with no separate titles outside
            picture_table = Table(
                [[old_box, new_box]],
                colWidths=[available_width / 2, available_width / 2]
            )
            picture_table.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('LEFTPADDING', (0,0), (-1,-1), 0),
                ('RIGHTPADDING', (0,0), (-1,-1), 0),
            ]))

            story.append(picture_table)
            story.append(Spacer(1, 25))





            # === Section 3: Change Plan and Break Point ===
            plan_text = next((col["text"] for col in item.get("column_values", []) 
                                if col["column"]["title"] == "Change Plan" and col.get("text")), "N/A")
            break_text = next((col["text"] for col in item.get("column_values", []) 
                                if col["column"]["title"] == "Break Point" and col.get("text")), "N/A")

            plan_para = Paragraph(f"<b>Change Plan</b><br/>{plan_text}", styleN)
            break_para = Paragraph(f"<b>Break Point</b><br/>{break_text}", styleN)


            # Wrap both paragraphs and tables into one red-bordered table box
            red_box = Table([[plan_para], [break_para]],
                colWidths=[available_width],
                style=TableStyle([
                    ('BOX', (0, 0), (-1, -1), 1, colors.red),
                    ('BACKGROUND', (0, 0), (-1, -1), colors.white),
                    ('LEFTPADDING', (0, 0), (-1, -1), 10),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 10),
                    ('TOPPADDING', (0, 0), (-1, -1), 8),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
                ])
            )

            story.append(red_box)
            story.append(Spacer(1, 20))




            # === Section 4: Actions of Each Department ===
            dept_actions = [
                ("Design", "BoM, Spec, Catalogue, Support"),
                ("Sales and After Sales", "Notification, Spare"),
                ("Marketing", "Update, Promotion"),
                ("Store", "Adjust, Excess"),
                ("Planning", "BoM, Spec, Schedule, Forecast"),
                ("Purchasing", "Supplier, Cost, Orders"),
                ("QC", "Standard, SOP"),
                ("Production", "Equipment, Training, Redvan"),
            ]

            action_paragraphs = []
            action_paragraphs.append(Paragraph("<b>Actions of Each Department</b>", styleB))

            for dept, tasks in dept_actions:
                bullet = f'<bullet>&bull;</bullet>'
                action_text = f'{bullet} <b>{dept}</b><br/>{tasks}'
                action_paragraphs.append(Paragraph(action_text, styleN))
                action_paragraphs.append(Spacer(1, 5))

            # Wrap into green-bordered box
            green_box = Table(
                [[action_paragraphs]],
                colWidths=[available_width],
                style=TableStyle([
                    ('BOX', (0, 0), (-1, -1), 1, colors.green),
                    ('LEFTPADDING', (0, 0), (-1, -1), 10),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 10),
                    ('TOPPADDING', (0, 0), (-1, -1), 8),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
                ])
            )

            story.append(green_box)
            story.append(PageBreak())


    doc.build(story, onFirstPage=draw_watermark, onLaterPages=draw_watermark)
    logger.info(f"PDF generated: {filename}")


# === Send Email via Outlook ===
def send_email_via_outlook(to, subject, body, attachment_path):
    outlook = win32.Dispatch("outlook.application")
    mail = outlook.CreateItem(0)
    mail.To = to
    mail.Subject = subject
    mail.Body = body
    mail.Attachments.Add(os.path.abspath(attachment_path))
    mail.Send()
    logger.info("Email sent successfully via Outlook.")

# === Save data for next run ===
def save_new_ids(new_started_items, previous_ids, filename):
    new_ids = [item["id"] for item in new_started_items if item["id"] not in previous_ids]

    if not new_ids:
        logger.info("No new items to save.")
        return

    updated_ids = sorted(previous_ids.union(new_ids))  # Optional: sorted

    try:
        with open(filename, "w", encoding="utf-8") as f:
            json.dump(updated_ids, f, indent=2)
        logger.info(f"Saved {len(new_ids)} new IDs to {filename}")
    except Exception as e:
        logger.error(f"Error saving data: {e}")


# === Main Execution ===
if __name__ == "__main__":
    try:
        logger.info("Loading previously saved item IDs...")
        previous_ids = load_previous_ids(PREVIOUS_DATA_FILE)

        logger.info("Fetching current 'Started' items from Monday.com...")
        new_started_items = get_board_items(previous_ids)

        if new_started_items:
            logger.info(f"Found {len(new_started_items)} new 'Started' items.")
            
            generate_pdf(new_started_items, PDF_OUTPUT_FILE)

            # send_email_via_outlook(
            #     to=EMAIL_TO,
            #     subject="ECO Board: New Project Started",
            #     body="Hi team, new project have been started. Please see attached report.",
            #     attachment_path=PDF_OUTPUT_FILE
            # )

            save_new_ids(new_started_items, previous_ids, PREVIOUS_DATA_FILE)

        else:
            logger.info("No new started project found.")

    except Exception as e:
        logger.error(f"Error occurred: {e}")

