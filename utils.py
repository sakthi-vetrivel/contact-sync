def confirm_action(message):
    """Prompt the user for confirmation (Y/n)."""
    return input(f"{message} (Y/n): ").strip().lower() == 'y'
