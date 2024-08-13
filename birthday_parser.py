from datetime import datetime
from dateutil import parser

def parse_birthday(date_str):
    """Parse a birthday string into a datetime object."""
    if isinstance(date_str, datetime):
        return date_str
    
    formats = ["%b-%d", "%B-%d", "%B %d", "%b %d", "%b %dth", "%B %dth", "%m/%d/%Y", "%m/%d/%y", "%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"]
    
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
