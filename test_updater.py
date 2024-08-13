import openpyxl
from AddressBook import (
    ABAddressBook, ABPerson, ABMutableMultiValue, 
    kABFirstNameProperty, kABLastNameProperty, kABPhoneProperty, 
    kABEmailProperty, kABBirthdayProperty, 
    kABPhoneHomeLabel, kABEmailHomeLabel
)
from datetime import datetime
from dateutil import parser  
import os

def create_test_contacts():
    """Create a set of unique test contacts in Apple Contacts."""
    address_book = ABAddressBook.sharedAddressBook()

    test_contacts = [
        {"first_name": "Xenon", "last_name": "Quasar", "phone": "123-456-7890", "email": "xenon.quasar@example.com", "birthday": "1990-01-01"},
        {"first_name": "Yara", "last_name": "Nova", "phone": "234-567-8901", "email": "yara.nova@example.com", "birthday": "Jan-15-1985"},
        {"first_name": "Zephyr", "last_name": "Lunar", "phone": "345-678-9012", "email": "zephyr.lunar@example.com", "birthday": None},
        {"first_name": "Aura", "last_name": "Solaris", "phone": None, "email": "aura.solaris@example.com", "birthday": "25 August 1992"},
    ]

    for contact in test_contacts:
        new_contact = ABPerson.alloc().init()
        new_contact.setValue_forProperty_(contact["first_name"], kABFirstNameProperty)
        new_contact.setValue_forProperty_(contact["last_name"], kABLastNameProperty)

        if contact["phone"]:
            phone_numbers = ABMutableMultiValue.alloc().init()
            phone_numbers.addValue_withLabel_(contact["phone"], kABPhoneHomeLabel)
            new_contact.setValue_forProperty_(phone_numbers, kABPhoneProperty)

        if contact["email"]:
            emails = ABMutableMultiValue.alloc().init()
            emails.addValue_withLabel_(contact["email"], kABEmailHomeLabel)
            new_contact.setValue_forProperty_(emails, kABEmailProperty)

        if contact["birthday"]:
            try:
                birthday_date = parser.parse(contact["birthday"])
                new_contact.setValue_forProperty_(birthday_date, kABBirthdayProperty)
            except ValueError:
                print(f"Could not parse birthday: {contact['birthday']} for {contact['first_name']} {contact['last_name']}")

        address_book.addRecord_(new_contact)

    address_book.save()
    print("Unique test contacts created successfully!")

def create_test_excel(file_path):
    """Create a test Excel file with diverse contact data."""
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    headers = ["First Name", "Last Name", "WhatsApp Number", "Personal Email", "Location after Graduation", "Social Media Handles", "Birthday"]
    sheet.append(headers)

    contacts_data = [
        # Similar names to test fuzzy matching and duplicates
        ["Xander", "Quasar", "123-456-7890", "xander.quasar@example.com", "New York", "@xander_quasar", "Jan-1"],
        ["Yara", "Nova", "234-567-8901", "yara.nova@example.com", "Los Angeles", "@yara_nova", "15-Jan-1985"],
        
        # Different last name, new contact
        ["Lyra", "Stellar", "456-789-0123", "lyra.stellar@example.com", "Boston", "@lyra_stellar", "25 Aug 1992"],
        
        # Existing contact, with slight variations to test updates
        ["Zephyr", "Lunar", "345-678-9012", "zephyr.lunar@example.com", "San Francisco", "@zephyr_lunar", None],
        ["Aura", "Solaris", "567-890-1234", "aura.solaris@example.com", "Chicago", "@aura_solaris", "August 25th 1992"],
        
        # Completely new contact with potential duplicates
        ["Nova", "Galaxy", "678-901-2345", "nova.galaxy@example.com", "Miami", "@nova_galaxy", "Sep-13"],
        ["Nova", "Galaxy", "678-901-2345", "nova.galaxy@example.com", "Miami", "@nova_galaxy2", "13 September 1993"]
    ]

    for contact in contacts_data:
        sheet.append(contact)

    workbook.save(file_path)
    print(f"Test Excel file created at {file_path}")

def delete_test_contacts():
    """Delete the test contacts from Apple Contacts."""
    address_book = ABAddressBook.sharedAddressBook()

    test_names = [
        {"first_name": "Xenon", "last_name": "Quasar"},
        {"first_name": "Yara", "last_name": "Nova"},
        {"first_name": "Zephyr", "last_name": "Lunar"},
        {"first_name": "Aura", "last_name": "Solaris"},
        {"first_name": "Xander", "last_name": "Quasar"},
        {"first_name": "Lyra", "last_name": "Stellar"},
        {"first_name": "Nova", "last_name": "Galaxy"},
    ]

    people = address_book.people()
    for test_name in test_names:
        for person in people:
            if (person.valueForProperty_(kABFirstNameProperty) == test_name["first_name"] and
                person.valueForProperty_(kABLastNameProperty) == test_name["last_name"]):
                address_book.removeRecord_(person)
                print(f"Deleted contact: {test_name['first_name']} {test_name['last_name']}")

    address_book.save()
    print("Unique test contacts deleted successfully!")

if __name__ == "__main__":
    # Define the path for the test Excel file
    test_excel_path = "test_contacts.xlsx"

    # Create a test Excel file with predefined data
    create_test_excel(test_excel_path)

    # Run the updater in different modes, cleaning up contacts between runs
    for mode, description in [
        ("", "default mode with a limit of 3"),
        ("--skeptical", "skeptical mode with a limit of 3"),
        ("", "full update with no limit")
    ]:
        print(f"\nRunning {description}:")
        
        # Create test contacts in Apple Contacts before each run
        create_test_contacts()
        
        # Run the updater script with the specified mode
        os.system(f"python3 updater.py {test_excel_path} {mode} --limit=3")

        # Clean up the contacts after each run
        delete_test_contacts()

    # Optionally, delete the test Excel file
    os.remove(test_excel_path)
    print(f"Deleted test Excel file: {test_excel_path}")
