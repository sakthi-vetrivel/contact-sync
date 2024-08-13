import re

def normalize_phone_number(phone_number):
    """Normalize phone number by removing non-digit characters and stripping country code for comparison."""
    phone_number = re.sub(r'\D', '', str(phone_number))
    
    if phone_number.startswith('1') and len(phone_number) > 10:
        phone_number = phone_number[1:]
    
    return phone_number

def format_phone_number_for_storage(phone_number, default_country_code="+1"):
    """Format phone number for storage, ensuring it includes the correct country code."""
    phone_number = str(phone_number)  # Ensure the phone number is a string
    normalized_number = normalize_phone_number(phone_number)
    return f"{default_country_code}{normalized_number}" if not phone_number.startswith('+') else phone_number
