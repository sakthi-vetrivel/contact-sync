import openpyxl
from datetime import datetime
from AddressBook import (
    ABAddressBook, ABPerson, ABMutableMultiValue,
    kABFirstNameProperty, kABLastNameProperty, kABPhoneProperty, 
    kABEmailProperty, kABPhoneHomeLabel, kABEmailHomeLabel, 
    kABBirthdayProperty
)
from phone_utils import normalize_phone_number, format_phone_number_for_storage
from birthday_parser import parse_birthday
from fuzzywuzzy import fuzz



def find_matching_contact(address_book, first_name, last_name, threshold=90):
    """Find a contact in the Address Book matching the given first and last name."""
    full_name = f"{first_name} {last_name}".strip().lower()
    matches = []

    for person in address_book.people():
        existing_first_name = (person.valueForProperty_(kABFirstNameProperty) or "").strip().lower()
        existing_last_name = (person.valueForProperty_(kABLastNameProperty) or "").strip().lower()
        existing_full_name = f"{existing_first_name} {existing_last_name}".strip()

        if full_name == existing_full_name or fuzz.ratio(full_name, existing_full_name) > threshold:
            matches.append(person)
    
    return matches

def add_or_update_contact_info(person, new_phone=None, new_email=None, new_birthday=None):
    """Add or update the contact's phone number, email, and birthday."""
    def add_or_update_phone_number():
        if new_phone:
            sanitized_new_number = normalize_phone_number(new_phone)
            phone_numbers = person.valueForProperty_(kABPhoneProperty)
            if phone_numbers:
                phone_numbers = phone_numbers.mutableCopy()
                existing_numbers = [normalize_phone_number(phone_numbers.valueAtIndex_(i)) for i in range(phone_numbers.count())]
                if sanitized_new_number not in existing_numbers:
                    formatted_number = format_phone_number_for_storage(new_phone)
                    phone_numbers.addValue_withLabel_(formatted_number, kABPhoneHomeLabel)
                    person.setValue_forProperty_(phone_numbers, kABPhoneProperty)
            else:
                formatted_number = format_phone_number_for_storage(new_phone)
                phone_numbers = ABMutableMultiValue.alloc().init()
                phone_numbers.addValue_withLabel_(formatted_number, kABPhoneHomeLabel)
                person.setValue_forProperty_(phone_numbers, kABPhoneProperty)

    def add_or_update_email():
        if new_email:
            emails = person.valueForProperty_(kABEmailProperty)
            if emails:
                emails = emails.mutableCopy()
                existing_emails = [emails.valueAtIndex_(i).strip().lower() for i in range(emails.count())]
                if new_email.strip().lower() not in existing_emails:
                    emails.addValue_withLabel_(new_email, kABEmailHomeLabel)
                    person.setValue_forProperty_(emails, kABEmailProperty)
            else:
                emails = ABMutableMultiValue.alloc().init()
                emails.addValue_withLabel_(new_email, kABEmailHomeLabel)
                person.setValue_forProperty_(emails, kABEmailProperty)

    def update_birthday():
        existing_birthday = person.valueForProperty_(kABBirthdayProperty)
        # If there is already a birthday, do not update it
        if existing_birthday or not(new_birthday):
            return
        
        birthday_date = parse_birthday(new_birthday)
        if birthday_date:
            # If the year is 2024, strip the year by setting it to a neutral year like 1900
            if birthday_date.year == 2024:
                birthday_date = birthday_date.replace(year=1900)

            # Set the birthday only if it's not already present
            person.setValue_forProperty_(birthday_date, kABBirthdayProperty)

    add_or_update_phone_number()
    add_or_update_email()
    update_birthday()

def generate_update_summary(person, new_phone, new_email, new_birthday):
    """Generate a summary of the updates to be made to the contact."""
    updates = []
    
    if new_phone:
        sanitized_new_phone = normalize_phone_number(new_phone)
        existing_phone_numbers = person.valueForProperty_(kABPhoneProperty)
        phone_numbers = [normalize_phone_number(existing_phone_numbers.valueAtIndex_(i)) for i in range(existing_phone_numbers.count())] if existing_phone_numbers else []
        if sanitized_new_phone not in phone_numbers:
            updates.append(f"Add phone number: {new_phone}")
    
    if new_email:
        existing_emails = person.valueForProperty_(kABEmailProperty)
        email_list = [existing_emails.valueAtIndex_(i).strip().lower() for i in range(existing_emails.count())] if existing_emails else []
        if new_email.strip().lower() not in email_list:
            updates.append(f"Add email: {new_email}")
    
    if new_birthday:
        existing_birthday = person.valueForProperty_(kABBirthdayProperty)
        if existing_birthday:
            existing_birthday = datetime.fromtimestamp(existing_birthday.timeIntervalSince1970())
        birthday_date = parse_birthday(new_birthday)
        if birthday_date:
            formatted_birthday = birthday_date.strftime('%m-%d') if birthday_date.year == 2024 else birthday_date.strftime('%Y-%m-%d')
            if not existing_birthday or (existing_birthday.date() != birthday_date.date()):
                updates.append(f"Update birthday: {formatted_birthday}")

    return "\n".join(updates) if updates else None

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
    return input(f"{message}\n(Y/n): ").strip().lower() == 'y'

def update_or_create_contact(address_book, contact_data, skeptical_mode):
    """Update an existing contact or create a new one based on the provided contact data."""
    first_name, last_name, whatsapp_number, personal_email, location, social_media_handles, birthday = contact_data[:7]
    full_name = f"{first_name} {last_name}".strip()
    
    normalized_phone = normalize_phone_number(whatsapp_number)
    
    # Find matching contacts
    matches = find_matching_contact(address_book, first_name, last_name, threshold=100)
    if not matches:
        matches = find_matching_contact(address_book, first_name, last_name, threshold=90)
    
    matched_person = None

    # Check for phone number match among the possible matches
    if matches:
        for person in matches:
            existing_phone_numbers = person.valueForProperty_(kABPhoneProperty)
            if existing_phone_numbers:
                phone_numbers = [normalize_phone_number(existing_phone_numbers.valueAtIndex_(i)) for i in range(existing_phone_numbers.count())]
                if normalized_phone in phone_numbers:
                    matched_person = person
                    break

    if skeptical_mode and not matched_person:
        matched_person = prompt_user_for_match(matches, full_name)
    elif len(matches) == 1 and not matched_person:
        matched_person = matches[0]
    
    if matched_person:
        proposed_updates = generate_update_summary(matched_person, whatsapp_number, personal_email, birthday)
        if proposed_updates and (not skeptical_mode or confirm_action(f"Update contact '{full_name}' with the following changes?\n{proposed_updates}")):
            add_or_update_contact_info(matched_person, whatsapp_number, personal_email, birthday)
            print(f"Updated contact: {full_name}")
    else:
        if not skeptical_mode or confirm_action(f"Create new contact '{full_name}'?"):
            new_contact = ABPerson.alloc().init()
            new_contact.setValue_forProperty_(first_name, kABFirstNameProperty)
            new_contact.setValue_forProperty_(last_name, kABLastNameProperty)
            add_or_update_contact_info(new_contact, whatsapp_number, personal_email, birthday)
            address_book.addRecord_(new_contact)
            print(f"Created new contact: {full_name}")

def process_contacts(file_path, skeptical_mode=False, limit=None):
    """Process the contacts from the Excel file and update or create them in the Address Book."""
    address_book = ABAddressBook.sharedAddressBook()
    
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    for i, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=1):
        if limit is not None and i > limit:
            break

        first_name = row[0]
        if not first_name:
            # Skip this row if the first name is missing
            continue
        
        update_or_create_contact(address_book, row, skeptical_mode)
    
    address_book.save()
    print("Contacts updated successfully!")
