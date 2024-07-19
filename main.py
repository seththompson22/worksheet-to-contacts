
# python-api@sheets-tutorial-428114.iam.gserviceaccount.com
import gspread
from google.oauth2.service_account import Credentials

scopes = [
    "https://www.googleapis.com/auth/spreadsheets"
]
creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
client = gspread.authorize(creds)

# gets the id of the sheet to access the document and saves the whole document under workbook
with open('sheets_contacts_test_id.txt', 'r') as file:
    sheet_id = file.read()
workbook = client.open_by_key(sheet_id)

# Sheet1 in this instance contains the contacts I want to access
sheet = workbook.worksheet("Sheet1")
# column one contains the titles of the fields I want to access for each possible new client
column_titles = sheet.row_values(1)

print(column_titles)

names = sheet.col_values(2)
print(names)


class Person:
    def __init__(self, name, role, email, phone_number):
        self.name = name
        self.role = role
        self.email = email
        self.phone_number = phone_number

contacts_from_sheet = []

for idx in range(len(names)-1):
    row = sheet.row_values(idx+1)
    name = row[1]
    role = row[2]
    email = row[3]
    phone_number = row[4]
    contacts_from_sheet.append(Person(name, role, email, phone_number))

for contact in contacts_from_sheet:
    print(f"{contact.name}, {contact.email}, {contact.phone_number}")

