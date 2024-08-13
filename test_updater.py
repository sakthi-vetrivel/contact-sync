import openpyxl
from datetime import datetime
from updater import update_contacts_from_excel  # Import the update function from the main updater script
from AddressBook import ABAddressBook, ABPerson, ABMutableMultiValue, kABFirstNameProperty, kABLastNameProperty, kABPhoneProperty, kABEmailProperty, kABPhoneHomeLabel, kABEmailHomeLabel, kABBirthdayProperty

def create_test_contacts():
    # Initialize Address Book
    address_book = ABAddressBook.sharedAddressBook()

    # Define test contacts with unique and similar names to trigger confirmation
    test_contacts = [
        {"first_name": "Xanther", "last_name": "Quizzik", "phone": "123-456-7890", "email": "xanther.quizzik@example.com", "birthday": "1990-01-01"},
        {"first_name": "Xylia", "last_name": "Quizzik", "phone": "234-567-8901", "email": "xylia.quizzik@example.com", "birthday": "1985-05-15"},
        {"first_name": "Orion", "last_name": "Stardust", "phone": "345-678-9012", "email": "orion.stardust@example.com", "birthday": None},
        {"first_name": "Eldric", "last_name": "Moonshadow", "phone": None, "email": "eldric.moonshadow@example.com", "birthday": "1992-08-25"},
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
            birthday_date = datetime.strptime(contact["birthday"], "%Y-%m-%d")
            new_contact.setValue_forProperty_(birthday_date, kABBirthdayProperty)

        address_book.addRecord_(new_contact)

    # Save changes to the Address Book
    address_book.save()
    print("Unique test contacts created successfully!")

def delete_test_contacts():
    # Initialize Address Book
    address_book = ABAddressBook.sharedAddressBook()

    # Define test contacts to delete with unique names
    test_names = [
        {"first_name": "Xanther", "last_name": "Quizzik"},
        {"first_name": "Xylia", "last_name": "Quizzik"},
        {"first_name": "Orion", "last_name": "Stardust"},
        {"first_name": "Eldric", "last_name": "Moonshadow"},
    ]

    # Search for and delete each contact
    people = address_book.people()
    for test_name in test_names:
        for person in people:
            if (person.valueForProperty_(kABFirstNameProperty) == test_name["first_name"] and
                person.valueForProperty_(kABLastNameProperty) == test_name["last_name"]):
                address_book.removeRecord_(person)
                print(f"Deleted contact: {test_name['first_name']} {test_name['last_name']}")

    # Save changes to the Address Book
    address_book.save()
    print("Unique test contacts deleted successfully!")

if __name__ == "__main__":
    # Run the test setup and update process
    create_test_contacts()

    # Specify the test Excel file (replace with the path to your actual test file)
    excel_file_path = "test_contacts.xlsx"
    update_contacts_from_excel(excel_file_path)

    # Prompt user to check the contacts
    input("Please check the updated contacts in your Address Book. Press Enter to continue and clean up test data...")

    # Clean up the test data
    delete_test_contacts()
