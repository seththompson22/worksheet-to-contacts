import time
import gspread
import os
import re
from google.oauth2.credentials import Credentials
from google.oauth2 import service_account
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly', 'https://www.googleapis.com/auth/contacts']
CLIENT_SECRETS_FILE = 'credentials.json'


# Person class to represent each contact
class Person:
    def __init__(self, name, email, phone_number):
        self.name = name
        self.email = email
        self.phone_number = phone_number


def read_spreadsheet():
    flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRETS_FILE, SCOPES)
    creds = flow.run_local_server(port=0)
    client = gspread.authorize(creds)

    # Get the ID of the sheet to access the document
    with open('sheets_contacts_test_id.txt', 'r') as file:
        sheet_id = file.read().strip()
    
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
    for idx in range(2, 52):
        row = sheet.row_values(idx)
        print(row)
        if len(row) < 5:  # Ensure row has at least 5 columns
            continue
        name = row[1]
        email = row[3]
        phone_number = row[4]
        
        contacts_from_sheet.append(Person(name, email, phone_number))

    return contacts_from_sheet


# Function to normalize phone number
def normalize_phone_number(phone_number):
    normalized_number = re.sub(r'\D', '', phone_number)  # Remove non-digit characters
    if len(normalized_number) > 10:
        normalized_number = normalized_number[-10:]  # Remove leading digits for internation numbers
    # add back in characters and digits as (123) 392-1343
    normalized_index = 0
    formatted_number = f"({normalized_number[:3]}) {normalized_number[3:6]}-{normalized_number[6:]}"
    
    return formatted_number

# Function to get authenticated credentials
def get_credentials():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        try:
            creds = Credentials.from_authorized_user_file('token.json')
        except ValueError:
            print("Could not access credentials from token.json")
            creds = None  # Handle error gracefully
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

def search_contact_by_criteria(service, contacts_to_search):    
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
    contacts_to_generate = []
    for i, person in enumerate(contacts_to_search):
        person_email = person.email.strip().lower()
        person_number = person.phone_number
        person_name = person.name.strip().lower()

        for connection in connections:
            connection_name = connection.get('names', [{}])[0].get('displayName', '').strip().lower()
            connection_emails = [email['value'].strip().lower() for email in connection.get('emailAddresses', [])]
            connection_phones = [phone['value'] for phone in connection.get('phoneNumbers', [])]

            is_contact_found = False

            # Check if the contact matches the criteria
            # matches_name = not name or name.lower() in person_name
            matches_email = person_email and person_email.strip().lower() in connection_emails
            # matches_phone = not phone or phone in person_phones

            if matches_email:
                found_contacts.append(connection)
                is_contact_found = True
                break

        if not is_contact_found:
            contacts_to_generate.append(Person(person_name, person_email, normalize_phone_number(person_number)))
            print(f"No contacts found matching the criteria. Name - {person_email}")

    if found_contacts:
        print('Contacts found:')
        for person in found_contacts:
            names = person.get('names', [])
            name = names[0].get('displayName', 'Unknown') if names else 'Unknown'

            email_addresses = person.get('emailAddresses', [])
            email = email_addresses[0].get('value', 'Unknown') if email_addresses else 'Unknown'


            phone_numbers = person.get('phoneNumbers', [])
            phone = phone_numbers[0].get('value', 'Unknown') if phone_numbers else 'Unknown'
            phone = normalize_phone_number(phone)
            print(f'{name} - {email} - {phone}')
    else:
        print('No matching contacts found.')

    return found_contacts, contacts_to_generate

def update_contacts(service, contacts_to_update, contacts_from_sheet):
    for contact in contacts_to_update:
        contact_resource_name = contact.get('resourceName')
        current_name = contact.get('names', [{}])[0].get('displayName', 'Unknown')
        current_email = contact.get('emailAddresses', [{}])[0].get('value', 'Unknown')
        current_phone_number = contact.get('phoneNumbers', [{}])[0].get('value', 'Unknown')
        contact_etag = contact.get('etag')  # Retrieve the etag

        if not contact_etag:
            print(f"Missing etag for contact {current_name} ({current_email}). Skipping update.")
            continue

        corresponding_contact = None
        for p in contacts_from_sheet:
            if p.email == current_email:
                corresponding_contact = p
                break
        
        if corresponding_contact:
            needs_update = False
            updated_info = {
                'etag': contact_etag  # Include the etag in the update request
            }

            if corresponding_contact.name.strip().lower() != current_name.strip().lower():
                needs_update = True
                updated_info['names'] = [{'displayName': corresponding_contact.name}]
            
            
            corresponding_contact.phone_number = normalize_phone_number(corresponding_contact.phone_number)
            if current_phone_number != 'Unknown':
                if corresponding_contact.phone_number != normalize_phone_number(current_phone_number):
                    needs_update = True
                    updated_info['phoneNumbers'] = [{'value': corresponding_contact.phone_number}]
        
            if needs_update:
                try:
                    service.people().updateContact(
                        resourceName=contact_resource_name,
                        updatePersonFields='names,phoneNumbers',
                        body=updated_info
                    ).execute()
                    print(f"Updated Contact {current_name} ({current_email})")
                except Exception as e:
                    print(f"Failed to update contact {current_name}: {e}")
            else:
                print(f"No updates needed for contact {current_name} ({current_email})")
        else:
            print(f"Corresponding contact for {current_email} not found in the sheet.")

# Function to create a new contacts
def generate_contacts(service, contacts):
    created_contacts = []

    for person in contacts:
        try:
            contact = {}

            if person.name:
                contact['names'] = [{'displayName': person.name}]
            
            if person.email:
                contact['emailAddresses'] = [{'value': person.email}]

            if person.phone_number:
                contact['phoneNumbers'] = [{'value': person.phone_number}]

            if contact:
                created_contact = service.people().createContact(body=contact).execute()
                created_contacts.append(created_contacts)
                print(f"Contact {person.name} created with resourceName: {created_contact.get('resourceName')}")
        except Exception as e:
            print(f"Failed to create contact for {person.name}: {e}")
    return created_contacts

# Main function to run the script
def main():
    creds = get_credentials()

    service = build('people', 'v1', credentials=creds)
    # warmup query
    response = service.people().searchContacts(
    query="",
    readMask="names,emailAddresses"
    ).execute()
    
    # Wait for a few seconds
    time.sleep(5)

    # two lists of people to modify with one sheet
    #   contacts_to_check: List of all contacts that might need updated pulled from google sheet
    #   where contacts_to_generate U B --- contacts_to_check

    # Read contacts from Google Sheet
    contacts_from_sheet = read_spreadsheet() # searched by email because of UDEL directory

    contacts_to_update, contacts_to_generate = search_contact_by_criteria(service, contacts_from_sheet)  
    # contacts_to_update: Subset of contacts_to_check => list of contacts that need updated names/numbers
    # contacts_to_generate: Subset of contacts_to_check => list of contacts that need to be created
    
    update_contacts(service, contacts_to_update, contacts_from_sheet)
    generate_contacts(service, contacts_to_generate)

    # print(f'Contact {person.name}, {person.email} already exists with resourceName: {existing_contact[0].get("resourceName")}')

if __name__ == '__main__':
    main()
