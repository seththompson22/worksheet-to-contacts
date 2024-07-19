
# python-api@sheets-tutorial-428114.iam.gserviceaccount.com
import gspread
from google.oauth2.service_account import Credentials

scopes = [
    "https://www.googleapis.com/auth/spreadsheets"
]
creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
client = gspread.authorize(creds)

with open('sheets_contacts_test_id.txt', 'r') as file:
    sheet_id = file.read()
workbook = client.open_by_key(sheet_id)


sheet = workbook.worksheet("Sheet1")
column_titles = sheet.row_values(1)

print(column_titles)
