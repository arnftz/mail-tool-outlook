# Automated Email Sender with Attachments Using Python and Outlook

This script provides a way to send emails in bulk using Microsoft Outlook and Python. It automates email creation, adding CC addresses, attaching files, and sending emails to recipients fetched from a CSV file.

## Features
- Sends emails to a list of recipients provided in a CSV file.
- Customizable email subject, HTML body, and signature.
- Supports attachments for each email.
- Includes a CC field to copy additional recipients.
- Error handling for missing attachments or email-sending issues.

---

## Prerequisites
1. **Microsoft Outlook** installed and configured on your system.
2. **Python 3.x** installed.
3. Install the required Python library:
   - `pywin32` (for interacting with Outlook).
   
   Install it using:
   ```bash
   pip install pywin32
