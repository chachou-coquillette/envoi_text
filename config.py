# Configuration for the telephonic message agent

# Path to the contacts CSV file (columns: name, phone)
CONTACTS_FILE = "contacts.csv"

# The message to send to all contacts
MESSAGE = "Bonjour, ceci est un message automatique pour tester mon script."

# Optional image path to attach before sending (empty string disables attachment)
IMAGE_FILE = r"C:\Users\Charlotte\3. Loisirs & Voyages\2026 Mariage\Save the date.gif"

# Delay between messages in seconds (to avoid being rate-limited)
DELAY_BETWEEN_MESSAGES = 5

# Delay in seconds to wait for the Phone Link app UI to respond
UI_TIMEOUT = 10
