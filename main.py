import gspread
import os
import re
from google.oauth2.credentials import Credentials
from google.oauth2 import service_account
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

# Person class to represent each contact
class Person:
    def __init__(self, name, phone_number, email):
        self.name = name
        self.phone_number = phone_number
        self.email = email

def read_spreadsheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets"
    ]
    creds = service_account.Credentials.from_service_account_file("credentials.json", scopes=scopes)
    client = gspread.authorize(creds)

    # Get the ID of the sheet to access the document
    with open('sheets_contacts_test_id.txt', 'r') as file:
        sheet_id = file.read()
    
    # Open the workbook by its ID
    workbook = client.open_by_key(sheet_id)

    # Access the Sheet1 in the workbook (assuming it contains the contacts)
    sheet = workbook.worksheet("Sheet1")
    
    # Fetch column titles (assuming row 1 contains headers)
    column_titles = sheet.row_values(1)
    print(column_titles)

    # Initialize list to store contacts from sheet
    contacts_from_sheet = []
    # Iterate through rows to fetch contacts
    for idx in range(1, 42):
        row = sheet.row_values(idx)
        print(row)
        name = row[1]
        email = row[3]
        phone_number = row[4]
        
        contacts_from_sheet.append(Person(name, phone_number, email))

    return contacts_from_sheet

# Define the scopes and credentials file path
SCOPES = ['https://www.googleapis.com/auth/contacts']
CLIENT_SECRETS_FILE = 'people_credentials.json'

# Function to normalize phone number
def normalize_phone_number(phone_number):
    normalized_number = re.sub(r'\D', '', phone_number)  # Remove non-digit characters
    if normalized_number.startswith('1'):
        normalized_number = normalized_number[1:]  # Remove leading '1' (country code for USA)
    return normalized_number

# Function to get authenticated credentials
def get_credentials():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json')
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                CLIENT_SECRETS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    return creds

def search_contact_by_criteria(service, name=None, email=None, phone=None):
    results = service.people().connections().list(
        resourceName='people/me',
        pageSize=1000,
        personFields='names,emailAddresses,phoneNumbers',
        sortOrder='FIRST_NAME_ASCENDING'
    ).execute()

    connections = results.get('connections', [])

    if not connections:
        print('No contacts found.')
        return []

    found_contacts = []
    for person in connections:
        person_name = person.get('names', [{}])[0].get('displayName', '').lower()
        person_emails = [email['value'].lower() for email in person.get('emailAddresses', [])]
        person_phones = [phone['value'] for phone in person.get('phoneNumbers', [])]

        # Check if the contact matches the name and optionally the email or phone
        if (name and name.lower() in person_name) and \
           ((not email or email.lower() in person_emails) or \
            (not phone or phone in person_phones)):
            found_contacts.append(person)

    if not found_contacts:
        print('No contacts found matching the criteria.')
    else:
        print('Contacts found:')
        for person in found_contacts:
            names = person.get('names', [])
            if names:
                name = names[0].get('displayName', 'Unknown')
            else:
                name = 'Unknown'

            email_addresses = person.get('emailAddresses', [])
            if email_addresses:
                email = email_addresses[0].get('value', 'Unknown')
            else:
                email = 'Unknown'

            phone_numbers = person.get('phoneNumbers', [])
            if phone_numbers:
                phone = phone_numbers[0].get('value', 'Unknown')
            else:
                phone = 'Unknown'

            print(f'{name} - {email} - {phone}')

    return found_contacts

# Function to create a new contact
def create_contact(service, person):
    contact = {
        'names': [{'displayName': person.name}],
        'phoneNumbers': [{'value': person.phone_number}],
        'emailAddresses': [{'value': person.email}]
    }

    created_contact = service.people().createContact(body=contact).execute()
    return created_contact

# Main function to run the script
def main():
    creds = get_credentials()

    service = build('people', 'v1', credentials=creds)

    # Read contacts from Google Sheet
    persons = read_spreadsheet()

    # Search or create contacts based on Google Sheet data
    for person in persons:
        existing_contact = search_contact_by_criteria(service, person.name, person.email, person.phone_number)
        if existing_contact:
            print(f'Contact {person.name} already exists with resourceName: {existing_contact[0].get("resourceName")}')
        else:
            print(f'Creating contact for: {person.name}')
            # Uncomment the line below to create the contact
            created_contact = create_contact(service, person)
            print(f'Contact {person.name} created with resourceName: {created_contact.get("resourceName")}')


if __name__ == '__main__':
    main()
