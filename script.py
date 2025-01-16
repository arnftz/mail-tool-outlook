import win32com.client
import os
import csv

def send_emails(email_list, subject, body_html, attachments=None, cc_address=None, signature_html=None):
    """
    Send emails to a list of recipients using Outlook.

    Parameters:
        email_list (list): List of recipient email addresses.
        subject (str): Subject of the email.
        body_html (str): HTML content for the email body.
        attachments (list): List of file paths to attach (default is None).
        cc_address (str): CC email address (default is None).
        signature_html (str): HTML content for the email signature (default is None).
    """
    try:
        # Initialize Outlook application
        outlook = win32com.client.Dispatch("Outlook.Application")
        print("Connected to Outlook")

        # Loop through each recipient
        for recipient in email_list:
            try:
                print(f"Preparing email for: {recipient}")
                
                # Create an email item
                mail = outlook.CreateItem(0)  # 0: Mail Item
                mail.To = recipient
                if cc_address:
                    mail.CC = cc_address
                mail.Subject = subject
                mail.HTMLBody = f"{body_html}<br><br>{signature_html or ''}"  # Add signature if provided

                # Add attachments if provided
                if attachments:
                    for attachment in attachments:
                        if os.path.exists(attachment):
                            mail.Attachments.Add(attachment)
                            print(f"Attachment added: {attachment}")
                        else:
                            print(f"Attachment not found: {attachment}")

                # Send the email
                mail.Send()
                print(f"Email successfully sent to {recipient}{' (CC: ' + cc_address + ')' if cc_address else ''}")

            except Exception as e:
                print(f"Error sending email to {recipient}: {e}")

    except Exception as e:
        print(f"An error occurred while initializing Outlook: {e}")


def load_email_list(file_path):
    """
    Load email addresses from a CSV file.

    Parameters:
        file_path (str): Path to the CSV file.

    Returns:
        list: List of email addresses.
    """
    email_list = []
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            csv_reader = csv.reader(f)
            email_list = [row[0].strip() for row in csv_reader if row]  # Exclude empty rows
        print(f"Loaded {len(email_list)} email addresses from {file_path}")
    except Exception as e:
        print(f"Error reading email list from {file_path}: {e}")
    return email_list


# Example usage
if __name__ == "__main__":
    # Path to the email list CSV file
    email_list_file = "mail_list.csv"

    # Load email addresses
    email_list = load_email_list(email_list_file)

    # Email details
    subject = "Your Subject Here"
    body_html = "<p>This is the email body.</p><p>Thank you!</p>"
    attachments = ["path/to/attachment1.pdf", "path/to/attachment2.png"]  # Replace with actual paths
    cc_address = "example.cc@domain.com"  # Replace with actual CC address
    signature_html = "<p>Best regards,</p><p>Your Name</p>"

    # Send emails
    send_emails(email_list, subject, body_html, attachments, cc_address, signature_html)
