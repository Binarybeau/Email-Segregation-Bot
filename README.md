# Email Segregation Bot using UiPath

This project is a Robotic Process Automation (RPA) solution built with **UiPath**. It automates the process of reading unread emails from Microsoft Outlook, extracting key metadata, and saving relevant attachments.

## 🚀 Features

- Reads **unread emails** from Microsoft Outlook.
- Extracts the following details:
  - 📧 Sender (From)
  - 📨 Recipient (To)
  - 📝 Subject
  - 🕒 Date & Time
  - 📂 File path of the email (optional)
- Writes all extracted details into an **Excel file**.
- 📎 **Handles email attachments**:
  - Identifies emails with attachments.
  - Downloads and saves each attachment to a folder on the **Desktop**.
  - Folder and file naming format:  
    **`<Subject>_<Sender>_<Date>`**
  - Ensures organized storage for quick access by stakeholders.

## 🛠 Requirements

- UiPath Studio (Community or Enterprise Edition)
- Microsoft Outlook desktop client (configured)
- Excel installed
- Windows OS (for desktop path automation)

## 🎯 Use Case

Originally developed as part of my internship, this bot was created to help streamline internal communications and document tracking by:

- Reducing manual work for segregating emails.
- Helping stakeholders access key data faster.
- Organizing attachments automatically for reference or processing.

##  Workflow

<img width="1024" height="1536" alt="ChatGPT Image Jul 14, 2025, 10_12_39 AM" src="https://github.com/user-attachments/assets/6f2e36ba-ad9a-4dbf-8924-7e31d1683e4a" />


## 🤝 Contribution

This is an individual project developed under guidance during my internship.  
Open to feedback or ideas for future enhancements (e.g., integration with SharePoint, Teams, or ticketing systems).

