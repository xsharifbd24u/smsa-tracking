import asyncio
import os
import re
from urllib.parse import urlencode
from datetime import datetime
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd
from playwright.async_api import async_playwright

input_file = "tracking_numbers.xlsx"
output_file = "SMSA_Tracking_Results.xlsx"

if not os.path.exists(input_file):
    print(f"❌ Error! File '{input_file}' not found!")
    exit()

try:
    df = pd.read_excel(input_file, dtype={"Tracking Number": str})
    tracking_numbers = df["Tracking Number"].dropna().astype(str).str.strip().tolist()
    print(f"✅ Found {len(tracking_numbers)} tracking numbers\n")
except Exception as e:
    print(f"❌ Excel Reading Error: {e}")
    exit()

async def batch_track():
    print("🔄 Starting tracking...\n")
    start_time = datetime.now()
    print(f"⏰ Started: {start_time.strftime('%Y-%m-%d %H:%M:%S')}\n")
    
    batch_size = 20
    all_results = []
    
    for batch_start in range(0, len(tracking_numbers), batch_size):
        batch = tracking_numbers[batch_start:batch_start + batch_size]
        print(f"\n📦 Batch {batch_start//batch_size + 1}: Processing {len(batch)} numbers...")
        
        query_params = {f"tracknumbers[{i}]": num for i, num in enumerate(batch)}
        target_url = "https://www.smsaexpress.com/trackingdetails?" + urlencode(query_params)
        
        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=True)
            context = await browser.new_context(
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
            )
            page = await context.new_page()
            
            try:
                print(f"🌐 Loading SMSA website (Batch {batch_start//batch_size + 1})...")
                await page.goto(target_url, wait_until="networkidle", timeout=60000)
                await page.wait_for_timeout(3000)
                
                page_text = await page.inner_text("body")
                
                for track_id in batch:
                    print(f"   📦 {track_id}")
                    
                    if track_id in page_text:
                        idx = page_text.find(track_id)
                        block = page_text[idx:idx + 1200]
                        
                        status = "In Transit"
                        status_match = re.search(r"Status\s*\n\s*([^\n]+(?:Delivered|Out for Delivery|In Transit|Shipment Received)[^\n]*)", block, re.IGNORECASE)
                        if status_match:
                            status = status_match.group(1).strip()
                        else:
                            delivered_match = re.search(r"Delivered,?\s*Signed by:\s*([^\n]+)", block, re.IGNORECASE)
                            if delivered_match:
                                status = f"Delivered, Signed by: {delivered_match.group(1).strip()}"
                            elif "Out for Delivery" in block:
                                status = "Out for Delivery"
                            elif "Shipment Received at SMSA Sorting Facility" in block:
                                status = "Shipment Received at SMSA Sorting Facility"
                            elif "In Transit" in block:
                                status = "In Transit"
                            else:
                                status_match2 = re.search(r"(?:Status|STATUS)\s*\n\s*([^\n]+)", block, re.IGNORECASE)
                                if status_match2:
                                    status = status_match2.group(1).strip()
                        
                        origin = "N/A"
                        from_match = re.search(r"From\s*\n\s*([^\n]+)", block, re.IGNORECASE)
                        if from_match:
                            origin = from_match.group(1).strip()
                            origin = re.sub(r'^From\s*', '', origin)
                            if "mé & Príncipe" in origin or "Sao Tome" in origin:
                                origin = "N/A"
                        
                        destination = "N/A"
                        to_match = re.search(r"To\s*\n\s*([^\n]+)", block, re.IGNORECASE)
                        if to_match:
                            destination = to_match.group(1).strip()
                            destination = re.sub(r'^To\s*', '', destination)
                            if "mé & Príncipe" in destination or "Sao Tome" in destination:
                                destination = "N/A"
                        
                        if origin == "N/A":
                            from_alt = re.search(r"From\s+([^\n]+)", block, re.IGNORECASE)
                            if from_alt:
                                origin = from_alt.group(1).strip()
                                if "mé & Príncipe" in origin:
                                    origin = "N/A"
                        
                        if destination == "N/A":
                            to_alt = re.search(r"To\s+([^\n]+)", block, re.IGNORECASE)
                            if to_alt:
                                destination = to_alt.group(1).strip()
                                if "mé & Príncipe" in destination:
                                    destination = "N/A"
                        
                        last_update = "N/A"
                        date_match = re.search(r"(\d{2}-[A-Za-z]{3}-\d{4})\s*\n\s*([A-Za-z]+)?", block, re.IGNORECASE)
                        if date_match:
                            last_update = date_match.group(1).strip()
                            if date_match.group(2):
                                last_update += f" ({date_match.group(2).strip()})"
                        
                        time_match = re.search(r"(\d{2}:\d{2}\s*[AP]M)", block, re.IGNORECASE)
                        if time_match and last_update != "N/A":
                            last_update += f" {time_match.group(1).strip()}"
                        
                        all_results.append({
                            "Tracking Number": track_id,
                            "Status": status,
                            "From": origin,
                            "To": destination,
                            "Last Update": last_update,
                            "Check Result": "✅ Success"
                        })
                        print(f"      ✅ Status: {status[:50]}...")
                        print(f"      📍 From: {origin} → To: {destination}")
                        
                    else:
                        all_results.append({
                            "Tracking Number": track_id,
                            "Status": "❌ Not Found",
                            "From": "N/A",
                            "To": "N/A",
                            "Last Update": "N/A",
                            "Check Result": "❌ Failed"
                        })
                        print(f"      ❌ Not Found")
                        
            except Exception as e:
                print(f"   ❌ Batch Processing Error: {e}")
                for track_id in batch:
                    all_results.append({
                        "Tracking Number": track_id,
                        "Status": "⚠️ Error",
                        "From": "N/A",
                        "To": "N/A",
                        "Last Update": "N/A",
                        "Check Result": "❌ Failed"
                    })
            
            await browser.close()
    
    end_time = datetime.now()
    duration = end_time - start_time
    print(f"\n⏰ Finished: {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"⏱️ Total Time: {duration.total_seconds():.2f} seconds")
    
    create_excel(all_results, start_time, end_time)

def create_excel(results, start_time, end_time):
    try:
        print("\n📊 Creating Excel file...")
        result_df = pd.DataFrame(results)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "SMSA Tracking Status"
        
        ws.merge_cells('A1:F1')
        title_cell = ws['A1']
        title_cell.value = "📦 SMSA EXPRESS TRACKING REPORT"
        title_cell.font = Font(name="Calibri", size=16, bold=True, color="1F497D")
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        
        ws.merge_cells('A2:F2')
        info_cell = ws['A2']
        info_cell.value = f"Generated: {start_time.strftime('%Y-%m-%d %H:%M:%S')} | Duration: {(end_time - start_time).total_seconds():.2f}s | Total Records: {len(results)}"
        info_cell.font = Font(name="Calibri", size=10, color="666666")
        info_cell.alignment = Alignment(horizontal="center", vertical="center")
        
        ws.merge_cells('A3:F3')
        
        headers = ["Tracking Number", "Status", "From", "To", "Last Update", "Check Result"]
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=4, column=col_idx)
            cell.value = header
        
        header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        row_alt_fill = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
        thin_border = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9"),
        )
        
        for row_idx, record in enumerate(results, start=5):
            ws.cell(row=row_idx, column=1, value=record["Tracking Number"])
            ws.cell(row=row_idx, column=2, value=record["Status"])
            ws.cell(row=row_idx, column=3, value=record["From"])
            ws.cell(row=row_idx, column=4, value=record["To"])
            ws.cell(row=row_idx, column=5, value=record["Last Update"])
            ws.cell(row=row_idx, column=6, value=record["Check Result"])
        
        for col_idx in range(1, 7):
            cell = ws.cell(row=4, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
        
        for row_idx in range(5, ws.max_row + 1):
            for col_idx in range(1, 7):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.border = thin_border
                if col_idx in [1, 5, 6]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                if row_idx % 2 == 0:
                    cell.fill = row_alt_fill
        
        column_widths = {'A': 20, 'B': 45, 'C': 25, 'D': 25, 'E': 20, 'F': 15}
        for col_letter, width in column_widths.items():
            ws.column_dimensions[col_letter].width = width
        
        footer_row = ws.max_row + 1
        ws.merge_cells(f'A{footer_row}:F{footer_row}')
        footer_cell = ws[f'A{footer_row}']
        footer_cell.value = f"Report generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | SMSA Express Tracking System"
        footer_cell.font = Font(name="Calibri", size=9, color="999999")
        footer_cell.alignment = Alignment(horizontal="center", vertical="center")
        
        wb.save(output_file)
        print(f"\n✅ Success! File saved: '{output_file}'")
        print(f"📁 Location: {os.path.abspath(output_file)}")
        print(f"📊 Total {len(results)} records processed")
        
        success_count = sum(1 for r in results if r["Check Result"] == "✅ Success")
        print(f"\n📊 Summary:")
        print(f"   ✅ Success: {success_count}")
        print(f"   ❌ Failed: {len(results) - success_count}")
        print(f"   ⏱️ Time: {(end_time - start_time).total_seconds():.2f} seconds")
        
    except PermissionError:
        print(f"\n❌ Error! File '{output_file}' is open in Excel.")
        print("👉 Please close the file and try again.")

if __name__ == "__main__":
    asyncio.run(batch_track())
