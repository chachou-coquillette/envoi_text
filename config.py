# Configuration for the telephonic message agent

# Path to the contacts CSV file (columns: name, phone)
CONTACTS_FILE = "contacts.csv"

# The message to send to all contacts
MESSAGE = (
	"Save the date!\n"
	"Florian et Charlotte se marient le 5 septembre 2026.\n"
	"En attendant merci de remplir ce formulaire qui nous aidera à organiser au mieux cet évènement: https://forms.gle/FiFfyRP9AyjQA21q7\n"
	"Nous serons ravis de vous avoir à nos côtés pour célébrer notre amour.\n"
	"Florian et Charlotte"
)

# Optional image path to attach before sending (empty string disables attachment)
IMAGE_FILE = r"C:\Users\Charlotte\3. Loisirs & Voyages\2026 Mariage\Save the date.mp4"

# Delay between messages in seconds (to avoid being rate-limited)
DELAY_BETWEEN_MESSAGES = 5

# Delay in seconds to wait for the Phone Link app UI to respond
UI_TIMEOUT = 10
