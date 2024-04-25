import imaplib
import email
import os

# Constants
IMAP_SETTINGS = {
    "SERVER": 'outlook.office365.com',
    "EMAIL": 'your_email@example.com',
    "PASSWORD": 'your_password',
    "SENDER_EMAIL": 'sender@example.com',
    "ROOT_FOLDER_PATH": '/path/to/root/folder/'
}

def connect_to_imap_server(settings):
    """Connect to the IMAP server."""
    mail = imaplib.IMAP4_SSL(settings["SERVER"])
    mail.login(settings["EMAIL"], settings["PASSWORD"])
    return mail

def search_emails_by_sender(mail, sender_email):
    """Search for emails from a specific sender."""
    mail.select('inbox')
    result, data = mail.search(None, 'FROM', sender_email)
    return data

def save_attachment_from_email(mail, email_id, root_folder_path):
    """Save attachment from an email to a folder with the matching target string."""
    result, data = mail.fetch(email_id, '(RFC822)')
    raw_email = data[0][1]
    msg = email.message_from_bytes(raw_email)

    for part in msg.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue

        attachment_data = part.get_payload(decode=True)
        filename = part.get_filename()
        if filename:
            # Iterate through subfolders in the root folder
            for folder_name in os.listdir(root_folder_path):
                folder_path = os.path.join(root_folder_path, folder_name)
                if os.path.isdir(folder_path):
                    # Check if the attachment contains the first 6 characters of the subfolder name
                    if folder_name[:6] in attachment_data.decode('utf-8'):
                        # Save attachment into the matching subfolder
                        filepath = os.path.join(folder_path, filename)
                        with open(filepath, 'wb') as f:
                            f.write(attachment_data)
                            print(f"Attachment '{filename}' saved to '{filepath}'")

def close_imap_connection(mail):
    """Close the IMAP connection."""
    mail.close()
    mail.logout()

# Connect to the IMAP server
mail = connect_to_imap_server(IMAP_SETTINGS)

# Search for emails from the specified sender
emails_data = search_emails_by_sender(mail, IMAP_SETTINGS["SENDER_EMAIL"])

# Iterate over the emails and save attachments
for num in emails_data[0].split():
    save_attachment_from_email(mail, num, IMAP_SETTINGS["ROOT_FOLDER_PATH"])

# Close the connection
close_imap_connection(mail)
