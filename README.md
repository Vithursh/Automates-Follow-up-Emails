# Automates Follow-up Emails

## Description

---

Automates Follow-up Emails is a Python-based automation tool that integrates with Microsoft Outlook to intelligently track and manage follow-up emails. The application automatically monitors flagged emails in your inbox and sends scheduled follow-up messages to contacts who haven't responded after a specified period. Using a CSV-based tracking system, it maintains a record of all flagged emails and automatically sends reminders at defined intervals (default: 7 days), helping you stay on top of important communications without manual intervention.

---

## Inspiration

---

The inspiration for this project came from the challenge of managing important email follow-ups in a busy professional environment. Many professionals flag emails requiring responses but often forget to send reminders when no reply is received. This tool was created to automate the tedious process of manually tracking which emails need follow-ups and when to send them. By leveraging Outlook's COM interface and Python's data processing capabilities, this project provides a hands-off solution that ensures no important email slips through the cracks.

## Table of Contents

- Features
- Prerequisites
- Installation

## Features

---

- **Automatic Email Scanning**: Continuously monitors your Outlook inbox for flagged emails
- **Intelligent Tracking**: Maintains a CSV database of all tracked emails and their status
- **Scheduled Follow-ups**: Automatically sends follow-up emails at configurable intervals (default: 7 days)
- **Duplicate Prevention**: Prevents duplicate tracking of the same email from the same sender
- **Auto-Unflagging**: Automatically removes the flag from emails once a follow-up is sent
- **Error Handling**: Robust error handling ensures the automation runs continuously without crashing
- **Configurable Parameters**: Easily customize follow-up intervals, email subject lines, and body text
- **Real-time Monitoring**: Checks for new flagged emails at regular intervals (default: 30 seconds)

## Prerequisites

Before you begin, ensure you have the following installed:

- **Python 3.7 or higher**
- **Microsoft Outlook** (installed and configured on your machine)
- **pip** (Python package manager)

## Installation

---

### Step-by-Step Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/Vithursh/Automates-Follow-up-Emails.git
   ```

2. **Navigate into the project directory:**
   ```bash
   cd Automates-Follow-up-Emails
   ```

3. **Install all required dependencies:**
   ```bash
   pip install -r requirements.txt
   ```
   
   This command will automatically install all third-party libraries needed to run the application, including:
   - `pywin32` - For Windows COM interface with Outlook
   - `pandas` - For CSV data handling and manipulation

4. **Run the application:**
   ```bash
   python main.py
   ```

The application will now start monitoring your Outlook inbox and automatically manage follow-up emails based on your configuration.

### Configuration (Optional)

To customize the application behavior, open `main.py` in your preferred text editor and modify these variables at the top of the file:

```python
CHECK_INTERVAL = 30        # Time in seconds between inbox scans
FOLLOWUP_DAYS = 7          # Number of days before sending a follow-up
SUBJECT_TEXT = "Follow up on the analysis report"  # Subject line for follow-ups
BODY_TEXT = "Hi,\n\njust following up on my previous message. Let me know when you get a chance.\n\nThanks,\nTrench Group"
```

### Troubleshooting

- **Outlook COM Error**: Ensure Outlook is installed and running on your system
- **Permission Issues**: Run Python as Administrator if you encounter permission errors when accessing Outlook
- **CSV File Issues**: If the tracking CSV becomes corrupted, simply delete `email_tracking.csv` and the application will create a fresh one on the next run