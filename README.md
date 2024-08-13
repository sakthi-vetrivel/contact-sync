# Contact Sync

This script helps you update your Apple Contacts using data from an Excel sheet. It automates the process of adding or updating contact details, ensuring that your contact information is accurate and up-to-date.

## Getting Started

### 1. Clone the Repository
Download the code to your computer:
```bash
git clone https://github.com/yourusername/contact-updater.git
```

### 2. Install Python
If you don't have Python installed, download it from [python.org](https://www.python.org/downloads/).

### 3. Install Required Libraries
Open your terminal and run:
```bash
pip install openpyxl python-dateutil
```

### 4. Prepare Your Excel File
Ensure your Excel file (.xlsx) has the following columns: `First Name`, `Last Name`, `WhatsApp Number`, `Personal Email`, `Location`, `Social Media Handles`, and `Birthday`.

### 5. Run the Script

The script accepts three main arguments:

- **Excel File Path**: (Required) The path to the `.xlsx` file that contains the contact data.
- **`--skeptical`**: (Optional) Enables Skeptical Mode, which prompts you for confirmation before making any changes or creating new contacts.
- **`--limit=n`**: (Optional) Limits the number of contacts to update, processing only the first `n` rows from the Excel file.

#### Example Usage

**Default Mode:**
```bash
python3 updater.py path_to_your_excel_file.xlsx --limit=5
```
- **Purpose**: Updates the first 5 contacts from your Excel file without asking for confirmation unless multiple matches are found.
- **Use Case**: Ideal for updating a small batch of contacts quickly.

**Skeptical Mode:**
```bash
python3 updater.py path_to_your_excel_file.xlsx --skeptical --limit=10
```
- **Purpose**: Asks for confirmation on each change, even if there's only one match.
- **Use Case**: Useful when dealing with sensitive data or when you want to manually verify each update.

**No Limit:**
```bash
python3 updater.py path_to_your_excel_file.xlsx
```
- **Purpose**: Updates all contacts in the Excel file without any limit.
- **Use Case**: Suitable for processing large datasets when you're confident in the data accuracy.

### Understanding the Arguments

- **Excel File Path**: The script needs a path to the Excel file (.xlsx) where your contact data is stored. Without this, the script won't run.
- **`--skeptical`**: This flag tells the script to enable Skeptical Mode, where every potential change is reviewed by you. It’s useful when you're unsure about the data and want to avoid errors.
- **`--limit=n`**: Use this argument to specify the number of contacts you want to update. For example, `--limit=10` processes only the first 10 rows. If you don’t provide this argument, the script will process all rows.

### Limitations of the Code

- **Apple Contacts Only**: This script is designed to work with Apple Contacts. It won't work with other contact management systems.
- **Excel Format**: The script only supports `.xlsx` files. Other file formats like `.csv` or `.xls` are not supported.
- **Limited Error Handling**: The script assumes that your Excel file is properly formatted. If the file has missing or incorrectly formatted data, the script may not function as expected.
- **No Duplicate Detection Across Files**: The script does not check for duplicates across multiple Excel files or previous runs. It only checks within the scope of a single run.

### Testing the Code

If you're new to the script or want to test it before updating real contacts, you can create a small test Excel file with a few dummy contacts. Use the `--limit` argument to ensure that only these test contacts are processed. After running the script, you can manually verify the changes in your Apple Contacts.

**Testing Example:**
```bash
python3 updater.py test_contacts.xlsx --limit=2 --skeptical
```
- **Purpose**: Updates only the first 2 contacts in the `test_contacts.xlsx` file, with you confirming each change.

This allows you to safely test the script and understand how it works before applying it to your entire contact list.

## Need Help?

If you encounter any issues, please open an issue on this GitHub repository.

