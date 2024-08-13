import openpyxl
from datetime import datetime
from dateutil import parser
from AddressBook import ABAddressBook, ABPerson, ABMutableMultiValue, kABFirstNameProperty, kABLastNameProperty, kABPhoneProperty, kABEmailProperty, kABPhoneHomeLabel, kABEmailHomeLabel, kABBirthdayProperty
import sys

def parse_birthday(date_str):
    """Parse a birthday string into a datetime object."""
    formats = [
        "%b-%d", "%B-%d", "%B %d", "%b %d", "%b %dth", "%B %dth", 
        "%m/%d/%Y", "%m/%d/%y", "%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"
    ]
    
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    
    try:
        return parser.parse(date_str)
    except ValueError:
        print(f"Could not parse date: {date_str}")
        return None

def find_matching_contact(address_book, last_name, first_name=None):
    """Find a contact in the Address Book matching the given last name (and first name if provided)."""
    matches = []
    for person in address_book.people():
        existing_first_name = person.valueForProperty_(kABFirstNameProperty) or ""
        existing_last_name = person.valueForProperty_(kABLastNameProperty) or ""
        
        if last_name == existing_last_name and (not first_name or first_name == existing_first_name):
            matches.append(person)
    
    return matches

def update_or_create_contact(address_book, contact_data, skeptical_mode):
    """Update an existing contact or create a new one based on the provided contact data."""
    first_name, last_name, whatsapp_number, personal_email, _, _, birthday = contact_data
    full_name = f"{first_name} {last_name}"
    
    matches = find_matching_contact(address_book, last_name, first_name if skeptical_mode else None)
    matched_person = None

    if skeptical_mode or len(matches) > 1:
        matched_person = prompt_user_for_match(matches, full_name)
    elif len(matches) == 1:
        matched_person = matches[0]
    
    if skeptical_mode and matched_person:
        if not confirm_action(f"Update contact '{full_name}' with new data?"):
            print(f"Skipped updating contact: {full_name}")
            return
    
    if matched_person:
        update_contact(matched_person, whatsapp_number, personal_email, birthday)
    else:
        if skeptical_mode and not confirm_action(f"Create new contact '{full_name}'?"):
            print(f"Skipped creating contact: {full_name}")
            return
        
        create_contact(address_book, contact_data)

def prompt_user_for_match(matches, full_name):
    """Prompt the user to confirm the match when multiple possible contacts are found."""
    for person in matches:
        existing_first_name = person.valueForProperty_(kABFirstNameProperty)
        existing_last_name = person.valueForProperty_(kABLastNameProperty)
        existing_full_name = f"{existing_first_name} {existing_last_name}"
        if confirm_action(f"Is '{full_name}' the same as '{existing_full_name}'?"):
            return person
    return None

def confirm_action(message):
    """Prompt the user for confirmation (Y/n)."""
    return input(f"{message} (Y/n): ").strip().lower() == 'y'

def update_contact(person, whatsapp_number, personal_email, birthday):
    """Update the contact with the given data."""
    if whatsapp_number:
        add_or_update_phone_number(person, whatsapp_number)
    
    if personal_email:
        add_or_update_email(person, personal_email)
    
    if birthday:
        update_birthday(person, birthday)
    
    print(f"Updated contact: {person.valueForProperty_(kABFirstNameProperty)} {person.valueForProperty_(kABLastNameProperty)}")

def create_contact(address_book, contact_data):
    """Create a new contact in the Address Book."""
    first_name, last_name, whatsapp_number, personal_email, _, _, birthday = contact_data
    
    new_contact = ABPerson.alloc().init()
    new_contact.setValue_forProperty_(first_name, kABFirstNameProperty)
    new_contact.setValue_forProperty_(last_name, kABLastNameProperty)
    
    if whatsapp_number:
        add_or_update_phone_number(new_contact, whatsapp_number)
    
    if personal_email:
        add_or_update_email(new_contact, personal_email)
    
    if birthday:
        update_birthday(new_contact, birthday)
    
    address_book.addRecord_(new_contact)
    print(f"Created new contact: {first_name} {last_name}")

def add_or_update_phone_number(person, new_number):
    """Add or update the contact's phone number if it's not already present."""
    phone_numbers = person.valueForProperty_(kABPhoneProperty)
    if phone_numbers:
        phone_numbers = phone_numbers.mutableCopy()
        existing_numbers = [phone_numbers.valueAtIndex_(i) for i in range(phone_numbers.count())]
        if new_number not in existing_numbers:
            phone_numbers.addValue_withLabel_(new_number, kABPhoneHomeLabel)
            person.setValue_forProperty_(phone_numbers, kABPhoneProperty)
    else:
        phone_numbers = ABMutableMultiValue.alloc().init()
        phone_numbers.addValue_withLabel_(new_number, kABPhoneHomeLabel)
        person.setValue_forProperty_(phone_numbers, kABPhoneProperty)

def add_or_update_email(person, new_email):
    """Add or update the contact's email if it's not already present."""
    emails = person.valueForProperty_(kABEmailProperty)
    if emails:
        emails = emails.mutableCopy()
        existing_emails = [emails.valueAtIndex_(i) for i in range(emails.count())]
        if new_email not in existing_emails:
            emails.addValue_withLabel_(new_email, kABEmailHomeLabel)
            person.setValue_forProperty_(emails, kABEmailProperty)
    else:
        emails = ABMutableMultiValue.alloc().init()
        emails.addValue_withLabel_(new_email, kABEmailHomeLabel)
        person.setValue_forProperty_(emails, kABEmailProperty)

def update_birthday(person, birthday):
    """Update the contact's birthday if it's different from the current one."""
    birthday_date = parse_birthday(birthday)
    if birthday_date:
        existing_birthday = person.valueForProperty_(kABBirthdayProperty)
        if existing_birthday != birthday_date:
            person.setValue_forProperty_(birthday_date, kABBirthdayProperty)

def process_contacts(file_path, skeptical_mode=False, limit=None):
    """Process the contacts from the Excel file and update or create them in the Address Book."""
    address_book = ABAddressBook.sharedAddressBook()
    
    # Load the Excel workbook and select the active sheet
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    # Iterate over the rows in the sheet, starting from the second row (to skip headers)
    for i, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=1):
        if limit is not None and i > limit:
            break
        
        update_or_create_contact(address_book, row, skeptical_mode)
    
    # Save the changes to the Address Book
    address_book.save()
    print("Contacts updated successfully!")

if __name__ == "__main__":
    # Parse command-line arguments
    skeptical_mode = "--skeptical" in sys.argv
    
    # Find and parse the --limit argument
    limit_arg = next((arg for arg in sys.argv if arg.startswith("--limit=")), None)
    limit = int(limit_arg.split("=")[1]) if limit_arg else None
    
    # Find and parse the Excel file path argument
    file_path_arg = next((arg for arg in sys.argv if arg.endswith(".xlsx")), None)
    if not file_path_arg:
        print("Error: Please provide the path to the Excel file as a command-line argument.")
        sys.exit(1)

    # Process contacts with the provided arguments
    process_contacts(file_path_arg, skeptical_mode=skeptical_mode, limit=limit)
