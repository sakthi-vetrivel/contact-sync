# Contact Sync Utility

This project is a utility tool to sync contact information from an Excel file into your macOS Address Book. It provides functionality for updating or creating contacts with details such as phone numbers, emails, and birthdays.

## tl;dr 

I'd test this out first by setting a limit of 10 contacts, using the --skeptical flag, then maybe 20, and then when you feel confident that it works pretty well, don't set a limit or the skeptical flag. 

 ```bash
    git clone https://github.com/yourusername/contact-sync.git
    cd contact-sync
    pip install -r requirements.txt

    python main.py path_to_your_excel_file.xlsx --skeptical --limit=10
 ```

## Table of Contents
- [Introduction](#introduction)
- [Features](#features)
- [Project Structure](#project-structure)
- [Installation](#installation)
- [Usage](#usage)
- [Testing](#testing)
- [License](#license)

## Introduction

The Contact Sync Utility simplifies the process of managing your contacts. It reads from an Excel file and either updates existing contacts or creates new ones based on the information provided. The utility also supports "skeptical mode," where it prompts you for confirmation before making changes to your Address Book.

## Features

- **Update or Create Contacts**: Sync phone numbers, emails, and birthdays from an Excel file.
- **Skeptical Mode**: Get prompted for confirmation before making any updates to your contacts.
- **Phone Number Normalization**: Automatically formats and normalizes phone numbers for consistency.
- **Birthday Parsing**: Handles various date formats for birthdays.

## Project Structure

```
contact_sync/
│
├── main.py                # Entry point for running the utility
├── contact_manager.py     # Core logic for finding, updating, and creating contacts
├── phone_utils.py         # Utilities for normalizing and formatting phone numbers
├── birthday_parser.py     # Functions for parsing and handling birthdays
├── utils.py               # General utility functions
└── test_contact_sync.py   # Unit tests for the project
```

### Files Description

- **`main.py`**: The main entry point for the script. It handles command-line arguments and calls the appropriate functions to process contacts.
- **`contact_manager.py`**: Contains the core logic for managing contacts, including finding existing contacts, generating update summaries, and adding or updating contact information.
- **`phone_utils.py`**: Handles phone number normalization and formatting to ensure consistency.
- **`birthday_parser.py`**: Provides functionality to parse birthday strings from different formats.
- **`utils.py`**: Contains helper functions, such as user confirmation prompts.
- **`test_contact_sync.py`**: Contains unit tests to verify the functionality of the project.

## Installation

### Prerequisites

- Python 3.6 or higher
- macOS with Address Book (Contacts) app

### Setup

1. **Clone the Repository**:

    ```bash
    git clone https://github.com/yourusername/contact-sync.git
    cd contact-sync
    ```

2. **Install Dependencies**:

    Use `pip` to install required dependencies.

    ```bash
    pip install -r requirements.txt
    ```

    If you don’t have a `requirements.txt` file yet, you can create it with the necessary dependencies:

    ```plaintext
    openpyxl
    python-dateutil
    fuzzywuzzy
    ```
   
   You may also need to install Apple's AddressBook module if it's not already installed.

## Usage

### Running the Script

1. **Basic Usage**:

    To run the script with a specified Excel file:

    ```bash
    python main.py path_to_your_excel_file.xlsx
    ```

2. **Skeptical Mode**:

    Use the `--skeptical` flag to enable skeptical mode, where the script will prompt you before making any changes:

    ```bash
    python main.py path_to_your_excel_file.xlsx --skeptical
    ```

3. **Limiting the Number of Contacts**:

    Use the `--limit` flag to limit the number of contacts processed:

    ```bash
    python main.py path_to_your_excel_file.xlsx --limit=10
    ```

### Excel File Format

The Excel file should have the following columns in the first row:

- `First Name`
- `Last Name`
- `WhatsApp Number`
- `Personal Email`
- `Location after Graduation`
- `Social Media Handles`
- `Birthday`

Ensure that the data in the columns follows these guidelines:

- **Phone Numbers**: Can be in any format (e.g., `(540) 226-2697`, `+1 (540) 226-2697`), and will be normalized by the script.
- **Emails**: Should be valid email addresses.
- **Birthdays**: Can be in multiple formats like `Sep-17`, `03/18/2024`, `March 5th`.

## Testing

To run the tests:

```bash
python -m unittest discover
```

This command will automatically discover and run the tests defined in the `test_contact_sync.py` file.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
