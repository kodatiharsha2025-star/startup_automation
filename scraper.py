import os
import requests
import xml.etree.ElementTree as ET
import pandas as pd
import openpyxl
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")

def scrape_product_hunt():
    leads = []
    url = "https://www.producthunt.com/feed"
    headers = {"User-Agent": "Mozilla/5.0"}
    try:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            root = ET.fromstring(response.content)
            namespace = {"atom": "http://www.w3.org/2005/Atom"}
            for entry in root.findall("atom:entry", namespace):
                title_elem = entry.find("atom:title", namespace)
                link_elem = entry.find("atom:link[@rel='alternate']", namespace)
                content_elem = entry.find("atom:content", namespace)
                published_elem = entry.find("atom:published", namespace)
                
                title = title_elem.text if title_elem is not None else "Unknown Startup"
                website = link_elem.attrib.get("href", "https://producthunt.com") if link_elem is not None else "https://producthunt.com"
                
                context = "Product Hunt tech launch needing explainer video or motion graphics."
                if content_elem is not None and content_elem.text:
                    import re
                    clean_text = re.sub(r'<[^>]+>', ' ', content_elem.text).strip()
                    clean_text = " ".join(clean_text.split())
                    if clean_text:
                        context = clean_text

                founded_year = datetime.now().year
                if published_elem is not None and published_elem.text:
                    try:
                        founded_year = int(published_elem.text[:4])
                    except:
                        pass

                domain = website.replace("https://", "").replace("http://", "").split("/")[0]
                email = f"contact@{domain}"

                leads.append({
                    "Company Name": title,
                    "Founded Year": founded_year,
                    "Website": website,
                    "Email": email,
                    "Phone": "Not Disclosed",
                    "Context": f"[Product Hunt] {context}"
                })
    except Exception as e:
        print("Error fetching Product Hunt feed:", e)
    return leads

def scrape_y_combinator():
    leads = []
    # Y Combinator public Algolia search endpoint for startup launches
    url = "https://hn.algolia.com/api/v1/search_by_date?tags=show_hn"
    headers = {"User-Agent": "Mozilla/5.0"}
    try:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            data = response.json()
            for hit in data.get("hits", [])[:15]:
                title = hit.get("title")
                url_site = hit.get("url")
                created_at = hit.get("created_at")
                
                if not title or not url_site:
                    continue
                
                founded_year = datetime.now().year
                if created_at:
                    try:
                        founded_year = int(created_at[:4])
                    except:
                        pass

                domain = url_site.replace("https://", "").replace("http://", "").split("/")[0]
                email = f"founders@{domain}"

                leads.append({
                    "Company Name": title.replace("Show HN: ", ""),
                    "Founded Year": founded_year,
                    "Website": url_site,
                    "Email": email,
                    "Phone": "Not Disclosed",
                    "Context": "Y Combinator / Show HN early-stage startup looking for high-retention product demo videos."
                })
    except Exception as e:
        print("Error fetching YC/HN feed:", e)
    return leads

def scrape_startup_leads():
    ph_leads = scrape_product_hunt()
    yc_leads = scrape_y_combinator()
    return yc_leads + ph_leads

def save_to_excel(leads, filename="startup_leads_vhglobals.xlsx"):
    df = pd.DataFrame(leads)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Startup Leads"
    ws.views.sheetView[0].showGridLines = True

    HEADER_FILL = PatternFill(start_color="1B2631", end_color="1B2631", fill_type="solid")
    HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    ZEBRA_FILL = PatternFill(start_color="F2F4F4", end_color="F2F4F4", fill_type="solid")
    REGULAR_FONT = Font(name="Calibri", size=11)
    BORDER_COLOR = Side(border_style="thin", color="BDC3C7")
    CELL_BORDER = Border(left=BORDER_COLOR, right=BORDER_COLOR, top=BORDER_COLOR, bottom=BORDER_COLOR)

    headers = list(df.columns)
    ws.append(headers)

    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row_idx, row_data in enumerate(df.values, start=2):
        ws.append(list(row_data))
        for col_idx in range(1, len(row_data) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = REGULAR_FONT
            cell.border = CELL_BORDER
            cell.alignment = Alignment(vertical="center")
            if row_idx % 2 == 1:
                cell.fill = ZEBRA_FILL

    ws.column_dimensions['A'].width = 24
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 35
    ws.column_dimensions['D'].width = 30
    ws.column_dimensions['E'].width = 18
    ws.column_dimensions['F'].width = 65

    wb.save(filename)
    return filename

def send_telegram_document(filename):
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        print("Telegram credentials missing.")
        return
    
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendDocument"
    with open(filename, "rb") as doc:
        files = {"document": doc}
        data = {"chat_id": TELEGRAM_CHAT_ID, "caption": "🚨 Live YC & Product Hunt Startup Leads for vhglobals"}
        response = requests.post(url, data=data, files=files)
        print("Telegram response:", response.status_code)

if __name__ == "__main__":
    leads = scrape_startup_leads()
    if leads:
        file_path = save_to_excel(leads)
        send_telegram_document(file_path)
    else:
        print("No leads fetched.")
