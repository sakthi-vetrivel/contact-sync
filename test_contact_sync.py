import unittest
from datetime import datetime
from contact_manager import find_matching_contact, add_or_update_contact_info, generate_update_summary, update_or_create_contact
from phone_utils import normalize_phone_number, format_phone_number_for_storage
from birthday_parser import parse_birthday
from utils import confirm_action

class TestContactSync(unittest.TestCase):

    def test_parse_birthday(self):
        """Test birthday parsing with various formats."""
        self.assertEqual(parse_birthday("Sep-17").strftime('%Y-%m-%d'), '2024-09-17')
        self.assertEqual(parse_birthday("03/18/2024").strftime('%Y-%m-%d'), '2024-03-18')
        self.assertEqual(parse_birthday("March 5th").strftime('%Y-%m-%d'), '2024-03-05')

    def test_normalize_phone_number(self):
        """Test phone number normalization."""
        self.assertEqual(normalize_phone_number("(540) 226-2697"), "5402262697")
        self.assertEqual(normalize_phone_number("+1 (540) 226-2697"), "5402262697")
        self.assertEqual(normalize_phone_number("1-540-226-2697"), "5402262697")

    def test_format_phone_number_for_storage(self):
        """Test phone number formatting for storage."""
        self.assertEqual(format_phone_number_for_storage("(540) 226-2697"), "+15402262697")
        self.assertEqual(format_phone_number_for_storage("+1 (540) 226-2697"), "+15402262697")
        self.assertEqual(format_phone_number_for_storage("540-226-2697"), "+15402262697")

    def test_find_matching_contact(self):
        """Test finding a matching contact."""
        address_book = self.mock_address_book()
        matches = find_matching_contact(address_book, "John", "Doe")
        self.assertEqual(len(matches), 1)
        self.assertEqual(matches[0].valueForProperty_(kABFirstNameProperty), "John")
        self.assertEqual(matches[0].valueForProperty_(kABLastNameProperty), "Doe")

    def test_add_or_update_contact_info(self):
        """Test adding or updating contact information."""
        address_book = self.mock_address_book()
        person = address_book.people()[0]
        add_or_update_contact_info(person, new_phone="540-226-2697", new_email="johndoe@example.com", new_birthday="Sep-17")
        self.assertEqual(person.valueForProperty_(kABPhoneProperty).valueAtIndex_(0), "+15402262697")
        self.assertEqual(person.valueForProperty_(kABEmailProperty).valueAtIndex_(0), "johndoe@example.com")
        self.assertEqual(person.valueForProperty_(kABBirthdayProperty).strftime('%Y-%m-%d'), '2024-09-17')

    def test_generate_update_summary(self):
        """Test generating update summaries."""
        address_book = self.mock_address_book()
        person = address_book.people()[0]
        summary = generate_update_summary(person, "540-226-2697", "johndoe@example.com", "Sep-17")
        self.assertIn("Add phone number: 540-226-2697", summary)
        self.assertIn("Add email: johndoe@example.com", summary)
        self.assertIn("Update birthday: 09-17", summary)

    def test_confirm_action(self):
        """Test user confirmation action."""
        # This can be tricky to test because it requires user input.
        # You might want to mock input() to simulate different user responses.
        pass

    def mock_address_book(self):
        """Create a mock address book for testing."""
        from AddressBook import ABAddressBook, ABPerson, kABFirstNameProperty, kABLastNameProperty
        address_book = ABAddressBook.sharedAddressBook()
        person = ABPerson.alloc().init()
        person.setValue_forProperty_("John", kABFirstNameProperty)
        person.setValue_forProperty_("Doe", kABLastNameProperty)
        address_book.addRecord_(person)
        return address_book

if __name__ == "__main__":
    unittest.main()
