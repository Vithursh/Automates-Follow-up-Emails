# Automates Follow-up Emails

## Description

---

Automates Follow-up Emails is a Python-based automation tool that integrates with Microsoft Outlook to intelligently track and manage follow-up emails sent from your account. The application automatically monitors flagged emails in your Outlook Sent Items folder and sends scheduled follow-up messages to recipients who haven't responded after a specified period. Using a CSV-based tracking system, it maintains a record of all flagged emails and automatically sends reminders at defined intervals (default: 1 hour), helping you stay on top of important communications without manual intervention. The tool also detects when recipients respond before the automated follow-up is sent and removes them from the tracking list.

---

## Inspiration

---

The inspiration for this project came from the challenge of managing important email follow-ups in a busy professional environment. Many professionals send important emails and flag them for follow-up, but often forget to send reminders when no reply is received within the expected timeframe. This tool was created to automate the tedious process of manually tracking which emails need follow-ups and when to send them. By leveraging Outlook's COM interface and Python's data processing capabilities, this project provides a hands-off solution that ensures no important email slips through the cracks while being smart enough to detect when a recipient has already responded.

## Table of Contents

- Features
- Prerequisites
- Installation

## Features

---

- **Sent Items Monitoring**: Monitors flagged emails in your Outlook Sent Items folder (not the inbox)
- **Intelligent Recipient Tracking**: Tracks recipients by email address and subject line for accurate follow-ups
- **Exchange User Support**: Handles both Exchange and SMTP email addresses correctly
- **Scheduled Follow-ups**: Automatically sends follow-up emails at configurable intervals (customizable: hours, minutes, or days)
- **Duplicate Prevention**: Prevents duplicate tracking by checking exact subject and sent date matches
- **Response Detection**: Intelligently detects when recipients respond before the automated follow-up is sent and stops the tracking cycle
- **Auto-Unflagging**: Automatically removes the flag from sent emails once a follow-up is processed
- **Error Handling**: Robust error handling ensures the automation runs continuously without crashing
- **Configurable Parameters**: Easily customize follow-up intervals, and body text
- **Real-time Monitoring**: Checks for flagged emails at regular intervals (default: 30 seconds)

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

The application will now start monitoring your Outlook Sent Items folder and automatically manage follow-up emails based on your configuration.

### Configuration (Optional)

To customize the application behavior, open `main.py` in your preferred text editor and modify these variables at the top of the file:

```python
CHECK_INTERVAL = 30              # Time in seconds between Sent Items scans
FOLLOWUP_TIME_TYPE = "days"      # Units of time (current default)
FOLLOWUP_TIME = 7                # Number of days before sending a follow-up
BODY_TEXT = "Hi,\n\njust following up on my previous message. Let me know when you get a chance.\n\nThanks,\nTrench Group"
```

### Customizing Follow-up Intervals

The follow-up interval options (hours, minutes, and days) are available for you to use based on your specific needs and use cases. You can switch between any of these intervals depending on how aggressive or lenient you want your follow-up strategy to be.

**Current Configuration:** `FOLLOWUP_TIME = 7` (follows up after 7 days)

**To change to a different interval:**
1. Change `FOLLOWUP_TIME_TYPE` either `days`, `hours` or `minutes`
2. Change `FOLLOWUP_TIME` to the amount of time based on the unit of time chosen

**Examples:**

- **For hourly follow-ups (aggressive):** Set `FOLLOWUP_TIME = 1` and `FOLLOWUP_TIME_TYPE = "hours"`
- **For daily follow-ups (standard business):** Set `FOLLOWUP_TIME = 1` and `FOLLOWUP_TIME_TYPE = "days"`
- **For quick follow-ups (very aggressive):** Set `FOLLOWUP_TIME = 1` and `FOLLOWUP_TIME_TYPE = "minutes"`

### How It Works

1. **Scanning**: The application scans your Outlook Sent Items folder every 30 seconds for flagged emails
2. **Tracking**: When a flagged email is found, it records the recipient, subject line, and sent date in `email_tracking.csv`
3. **Follow-up Scheduling**: The tool schedules the next follow-up based on your configured interval (hours, minutes, or days)
4. **Sending Follow-ups**: When the scheduled time arrives, an automated follow-up email is sent to the recipients
5. **Response Detection**: If the recipients respond before the automated follow-up is sent, the email is automatically removed from tracking
6. **No Response**: If the recipients do not respond, the script will keep sending the automated email based on the `FOLLOWUP_TIME_TYPE` and `FOLLOWUP_TIME`
7. **Auto-Unflagging**: After processing, the flag is removed from the original sent email

### Troubleshooting

- **Outlook COM Error**: Ensure Outlook is installed and running on your system
- **Permission Issues**: Run Python as Administrator if you encounter permission errors when accessing Outlook
- **CSV File Issues**: If the tracking CSV becomes corrupted, simply delete `email_tracking.csv` and the application will create a fresh one on the next run, but keep in mind any currently tracked emails will get reset.
- **Exchange vs SMTP**: The application automatically detects whether recipients use Exchange or SMTP email addresses and handles them accordingly
- **No Emails Being Tracked**: Ensure you're flagging emails in your Sent Items folder, not the inbox
- **Follow-ups Not Triggering**: Verify that you've updated both the configuration variable at the top AND both code locations (scan_flagged_emails and process_followups functions) when changing the interval type
- **Setting Different `FOLLOWUP_TIME_TYPE` Or `FOLLOWUP_TIME`**: Whenever you are changing `FOLLOWUP_TIME_TYPE` Or `FOLLOWUP_TIME` for an exiting flag email or a new email, you must restart the script each time
- **Invalid FOLLOWUP_TIME_TYPE Error**: If your terminal displays a message like `ValueError: Invalid FOLLOWUP_TIME_TYPE: 'day'. Must be 'days', 'hours', or 'minutes'.` it means you have not typed the `FOLLOWUP_TIME_TYPE` incorrectly. It has to either be spelled like `days`, `hours` or `minutes` exactly
