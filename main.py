import sys
from contact_manager import process_contacts

def main():
    """Main function to handle command-line arguments and run the contact sync process."""
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

if __name__ == "__main__":
    main()
