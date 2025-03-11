# Email Automation with Python

This repository contains a simple Python script to automate the sending of emails using PowerShell and Microsoft Outlook APP. It is designed to make sending emails through Outlook easier with just a few lines of code.

## Features

- Sends emails via Microsoft Outlook App.
- Allows setting custom recipients (`To` and `CC`), subject, and body of the email.
- Uses Python’s `subprocess` module to invoke PowerShell commands for creating and sending emails.
- Simple to customize and use for quick email automation tasks.

## Requirements

Before running the script, ensure you have the following installed:

- Python 3.x
- Microsoft Outlook
- PowerShell (for running the automation commands)

## How It Works

The Python script uses the `subprocess` module to invoke a PowerShell script, which interacts with Outlook's COM object to create and send emails.

### Workflow

1. The script loops through a list of email addresses.
2. For each address, it sets up the email’s `To`, `CC`, subject, and body.
3. The script then uses PowerShell to create an email item in Outlook and sends it.
4. After sending the email, the script waits for a few seconds before continuing.

## How to Use

1. Clone or download this repository to your local machine.

    ```bash
    git clone https://github.com/yourusername/email-automation-python.git
    cd email-automation-python
    ```

2. Edit the script (`email_automation.py`) and update the following:
   - `addresses`: Add the list of email addresses you want to send emails to.
   - `email_subject`: Set the subject of the email.
   - `email_body`: Set the body of the email.

3. Run the script.

    ```bash
    python email_automation.py
    ```

The script will send the email(s) to the specified recipients via Outlook app.


