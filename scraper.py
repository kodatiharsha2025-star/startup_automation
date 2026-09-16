import os
import requests
import time
import re
import xml.etree.ElementTree as ET
from bs4 import BeautifulSoup
import pandas as pd
import openpyxl
from datetime import datetime, timedelta
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")

def clean_domain(url):
    clean = url.replace("https://", "").replace("http://", "").split("/")[0]
    clean = clean.replace("www.", "")
    return clean

def extract_real_email_from_website(url):
    # Attempts to find a real published contact email from the homepage
    try:
        headers = {"User-Agent": "Mozilla/5.0"}
        res = requests.get(url, headers=headers, timeout=4)
        if res.status_code == 200:
            soup = BeautifulSoup(res.text, 'html.parser')
            # Look for mailto links
            mailtos = soup.find_all('a', href=re.compile(r'^mailto:'))
            for m in mailtos:
                email = m['href'].replace('mailto:', '').split('?')[0].strip()
                if '@' in email and '.' in email:
                    return email
            
            # Look via regex for email pattern in text
            emails = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', res.text)
            for e in emails:
                if not any(x in e.lower() for x in ['example', 'png', 'jpg', 'w3.org', 'domain', 'sentry']):
                    return e
    except:
        pass
    
    # Fallback to smart pattern if no public email found on homepage
    domain = clean_domain(url)
    return f"founders@{domain}"

def scrape_historical_y_combinator():
    leads = []
    # Calculate timestamp for 6 months ago
    six_months_ago = datetime.now() - timedelta(days=180)
    timestamp_limit = int(six_months_ago.timestamp())
    
    # Loop through multiple pages to capture 6 months of historical Show HN launches
    for page in range(10): # Pulls up to 1000 historical records across pages
        url = f"https://hn.algolia.com/api/v1/search_by_date?tags=show_hn&numericFilters=created_at_i>{timestamp_limit}&hitsPerPage=100&page={page}"
        headers = {"User-Agent": "Mozilla/5.0"}
        try:
            response = requests.get(url, headers=headers, timeout=15)
            if response.status_code == 200:
                data = response.json()
                hits = data.get("hits", [])
                if not hits:
                    break
                
                for hit in hits:
                    title = hit.get("title")
                    url_site = hit.get("url")
                    created_at = hit.get("created_at")
                    
                    if not title or not url_site or "github.com" in url_site:
                        continue

                    founded_year = datetime.now().year
                    if created_at:
                        try:
                            founded_year = int(created_at[:4])
                        except:
                            pass

                    # Try to extract real email or use clean fallback
                    email = extract_real_email_from_website(url_site)

                    leads.append({
                        "Company Name": title.replace("Show HN: ", ""),
                        "Founded Year": founded_year,
                        "Website": url_site,
                        "Email": email,
                        "Phone": "Not Disclosed",
                        "Context": "Historical Show HN / YC startup from the past 6 months looking for explainer & demo videos."
                    })
                time.sleep(0.5) # Be polite to the free API endpoint
        except Exception as e:
            print(f"Error fetching historical page {page}:", e)
            break
            
    return leads

def save_to_excel(leads, filename="startup_leads_vhglobals.xlsx"):
    df = pd.DataFrame(leads).drop_duplicates(subset=["Website"])
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Startup Leads 6M"
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

    ws.column_dimensions['A'].width = 28
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 38
    ws.column_dimensions['D'].width = 32
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
        data = {"chat_id": TELEGRAM_CHAT_ID, "caption": "🚨 Full 6-Month Historical Startup Dataset for vhglobals"}
        response = requests.post(url, data=data, files=files)
        print("Telegram response:", response.status_code)

if __name__ == "__main__":
    leads = scrape_historical_y_combinator()
    if leads:
        file_path = save_to_excel(leads)
        send_telegram_document(file_path)
    else:
        print("No historical leads fetched.")
