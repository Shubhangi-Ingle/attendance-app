from fastapi import FastAPI, Form, File, UploadFile
from fastapi.responses import FileResponse, JSONResponse
import openpyxl
from openpyxl import Workbook
import os
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials

app = FastAPI()

EXCEL_FILE = "attendance.xlsx"

# ---- Google Sheets setup ----
GOOGLE_SHEET_ID = "1EwFrEzfHCDy1rNNRGBhWoFV4_Gxfw8D8p360eTBvDhU"
CREDENTIALS_FILE = "credentials.json"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

def get_google_sheet():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(GOOGLE_SHEET_ID).sheet1
    return sheet

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
    # Server-side date & time (trustworthy, not from student's phone)
    now = datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%H:%M:%S")

    # Photo is required (proves the student took a live picture) but we do NOT
    # save it anywhere — just read it to confirm it's a real, non-empty capture.
    content = await photo.read()
    if not content:
        return JSONResponse({"message": "No photo captured"}, status_code=400)

    # Open or create the Excel file
    if os.path.exists(EXCEL_FILE):
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.append(["Name", "Subject", "Date", "Time"])

    # Add new row
    ws.append([name, subject, date_str, time_str])

    wb.save(EXCEL_FILE)

    # Also write the same row into the live Google Sheet
    try:
        sheet = get_google_sheet()
        if len(sheet.get_all_values()) == 0:
            sheet.append_row(["Name", "Subject", "Date", "Time"])
        sheet.append_row([name, subject, date_str, time_str])
    except Exception as e:
        print("Google Sheet write failed:", e)

    return JSONResponse({"message": "Attendance marked successfully"})