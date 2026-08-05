from fastapi import FastAPI, Form, File, UploadFile
from fastapi.responses import FileResponse, JSONResponse
import openpyxl
from openpyxl import Workbook
import os
from datetime import datetime, timezone, timedelta
import gspread
import cloudinary
import cloudinary.uploader
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import io

app = FastAPI()

EXCEL_FILE = "attendance.xlsx"
IST = timezone(timedelta(hours=5, minutes=30))

# ---- Google Sheets & Drive setup ----
GOOGLE_SHEET_ID = "1EwFrEzfHCDy1rNNRGBhWoFV4_Gxfw8D8p360eTBvDhU"
CREDENTIALS_FILE = "credentials.json"

# ---- Cloudinary setup ----
CLOUDINARY_CLOUD_NAME = "do2oynvcf"
CLOUDINARY_API_KEY = "124479454635314"
CLOUDINARY_API_SECRET = "3kxzud_8JY1b53A8Jm2tEC1qjHE"

cloudinary.config(
    cloud_name=CLOUDINARY_CLOUD_NAME,
    api_key=CLOUDINARY_API_KEY,
    api_secret=CLOUDINARY_API_SECRET
)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

def get_google_sheet():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(GOOGLE_SHEET_ID).sheet1
    return sheet

def upload_photo_to_cloudinary(content: bytes, filename: str) -> str:
    """Uploads photo bytes to Cloudinary, returns the secure URL."""
    try:
        upload_result = cloudinary.uploader.upload(io.BytesIO(content), public_id=os.path.splitext(filename)[0])
        return upload_result["secure_url"]
    except Exception as e:
        print(f"Cloudinary upload failed: {e}")
        return "Upload failed"

# Serve the homepage (index.html)
@app.get("/")
def serve_form():
    return FileResponse("index.html")

# Let the teacher download the latest attendance Excel file anytime
@app.get("/download-excel")
def download_excel():
    return FileResponse(
        path=EXCEL_FILE,
        filename="attendance.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.post("/submit")
async def submit_attendance(
    name: str = Form(...),
    subject: str = Form(...),
    photo: UploadFile = File(...)
):
    # Use IST regardless of the server's own timezone (Render runs in UTC)
    now = datetime.now(IST)
    date_str = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%H:%M:%S")

    # Photo is required (proves the student took a live picture)
    content = await photo.read()
    if not content:
        return JSONResponse({"message": "No photo captured"}, status_code=400)

    # Upload the photo to Google Drive and get a viewable link
    safe_name = name.strip().replace(" ", "_")
    timestamp_str = now.strftime("%Y%m%d_%H%M%S")
    drive_filename = f"{safe_name}_{subject}_{timestamp_str}.jpg"

    try:
        photo_link = upload_photo_to_cloudinary(content, drive_filename)
    except Exception as e:
        print("Drive upload failed:", e)
        photo_link = "Upload failed"

    # Open or create the Excel file
    if os.path.exists(EXCEL_FILE):
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.append(["Name", "Subject", "Date", "Time", "Photo"])

    # Add new row (display "View Photo" if upload succeeded)
    display_text = "View Photo" if photo_link != "Upload failed" else photo_link
    ws.append([name, subject, date_str, time_str, display_text])

    # Make the photo link clickable in Excel
    last_row = ws.max_row
    cell = ws.cell(row=last_row, column=5)
    if photo_link != "Upload failed":
        cell.hyperlink = photo_link
        cell.style = "Hyperlink"
        cell.font = openpyxl.styles.Font(color="0000FF", underline="single")

    wb.save(EXCEL_FILE)

    # Also write the same row into the live Google Sheet
    try:
        sheet = get_google_sheet()
        # Check if the sheet is empty or only contains headers
        all_values = sheet.get_all_values()
        if not all_values or (len(all_values) == 1 and all_values[0] == ["Name", "Subject", "Date", "Time", "Photo"]):
            sheet.append_row(["Name", "Subject", "Date", "Time", "Photo"])
        # Use HYPERLINK formula for Google Sheets to show "View Photo"
        display_link = f'=HYPERLINK("{photo_link}", "View Photo")' if photo_link != "Upload failed" else photo_link
        sheet.append_row([name, subject, date_str, time_str, display_link], value_input_option="USER_ENTERED")
    except Exception as e:
        print("Google Sheet write failed:", e)

    return JSONResponse({"message": "Attendance marked successfully"})