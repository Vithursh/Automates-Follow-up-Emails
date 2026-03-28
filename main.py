import win32com.client
import pandas as pd
import os
import time
import pytz
import time
from datetime import datetime, timedelta, timezone

CSV_FILE = "email_tracking.csv"
CHECK_INTERVAL = 30  # seconds
FOLLOWUP_DAYS = 7 # 7 days until next follow-up if no response

SUBJECT_TEXT = "Follow up on the analysis report"
BODY_TEXT = "Hi,\n\njust following up on my previous message. Let me know when you get a chance.\n\nThanks,\nTrench Group"

# -------------------------------
# Initialize CSV if not exists
# -------------------------------
def init_csv():
    if not os.path.exists(CSV_FILE):
        df = pd.DataFrame(columns=[
            "email",
            "subject_line",
            "sent_date",
            "flagged_date",
            "last_seen_date",
            "next_followup_due"
        ], dtype="object")
        df.to_csv(CSV_FILE, index=False)

# -------------------------------
# Load CSV
# -------------------------------
def load_data():
    return pd.read_csv(CSV_FILE, parse_dates=[
        "sent_date", "flagged_date", "last_seen_date", "next_followup_due"
    ])

# -------------------------------
# Save CSV
# -------------------------------
def save_data(df):
    df.to_csv(CSV_FILE, index=False)

# -------------------------------
# Get Outlook Inbox
# -------------------------------
def get_outlook_inbox():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox
    return inbox


# -------------------------------
# Scan flagged emails today
# -------------------------------
def scan_flagged_emails(df):
    print("Scanning for flagged emails...")
    
    # Get the Outlook inbox folder
    inbox = get_outlook_inbox()
    
    # Get all messages (emails) from the inbox
    messages = inbox.Items
    print(f"Total emails in inbox: {len(messages)}")
    
    # Sort emails by received time in descending order (newest first)
    messages.Sort("[ReceivedTime]", True)

    # Loop through each email in the inbox
    for msg in messages:
        print("The loaded dataframe is:", df)
        # print("The emails currently being tracked are:", tracked_emails["email"].tolist())
        # Print debug info for each message
        print("Messages today:", msg)
        
        try:
            # Get the time the email was received
            received_time = msg.ReceivedTime

            # Get the sender's email address and convert to lowercase
            sender = msg.SenderEmailAddress.lower()
            print(f"Scanning email from: {sender}")  # Print who we're checking

            # Check if this email has been flagged for follow-up (FlagStatus = 2 means flagged)
            if msg.FlagStatus == 2:  # 2 = flagged
                # Get the sender again (already done above, but kept for clarity)
                sender = msg.SenderEmailAddress.lower()

                # Option A: Make everything timezone-aware (UTC)
                now = datetime.now().replace(tzinfo=None)
                # 2. Force 'received_time' to be naive
                if received_time.tzinfo is not None:
                    received_time = received_time.replace(tzinfo=None)

                # Check if this sender does not exist in our tracking CSV
                print(f"Email from CSV file: {df['email'].tolist()}")  # Print current tracked emails
                if not (df["email"] == sender).any():
                    # If they don't exist, create a new row with their info
                    new_row = {
                        "email": sender,
                        "subject_line": msg.Subject,
                        "sent_date": received_time,
                        "flagged_date": now,
                        "last_seen_date": now,
                        "next_followup_due": now + timedelta(days=FOLLOWUP_DAYS)
                    }

                    # 2. DO NOT use .astype(object). Let pandas keep native datetime dtypes.
                    new_entry = pd.DataFrame([new_row])  # Create a DataFrame for the new row

                    # Create a list of DataFrames to join
                    to_concat = [df, new_entry]

                    # Only include DataFrames that are not empty
                    valid_dfs = [d for d in to_concat if not d.empty]

                    if valid_dfs:
                        df = pd.concat(valid_dfs, ignore_index=True)
                    else:
                        # If both were empty, you can just keep the original df or create a fresh one
                        df = df 
            else:
                # If the sender exists, and they are unflagged in inbox, remove them from tracking
                if df["email"].isin([sender]).any() and df["subject_line"].isin([msg.Subject]).any():
                    # remove from tracking if unflagged
                    df = df[df["email"] != sender]
                    print(f"Removed {sender} from tracking because it was unflagged")

        # If any error occurs while processing an email, print the error and continue
        except Exception as e:
            print(f"Error processing email: {e}")

    # Return the updated dataframe with all flagged emails
    return df

# -------------------------------
# Send follow-up email
# -------------------------------
def send_email(to_address, subject_line):
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    mail.To = to_address
    mail.Subject = SUBJECT_TEXT
    mail.Body = BODY_TEXT

    mail.Send()
    print(f"Follow-up sent to {to_address}")

    print("Email has been sent, going to unflagged the email in the inbox...")
    time.sleep(30)  # Wait a few seconds to ensure the email is sent before checking the inbox again

    # Get the Outlook inbox folder
    inbox = get_outlook_inbox()
    
    # Get all messages (emails) from the inbox
    messages = inbox.Items
    
    # Sort emails by received time in descending order (newest first)
    messages.Sort("[ReceivedTime]", True)

    # Loop through each email in the inbox
    for msg in messages:
        try:
            sender = msg.SenderEmailAddress.lower()
            if sender == to_address and msg.Subject == subject_line:
                print(f"Found email from {to_address} with subject {subject_line} for follow-up.")
                # Check if it's a standard email (MailItem)
                if msg.Class == 43: # 43 is the constant for olMail
                    print("Unflagging the email...")
                    
                    # This one line removes the flag, the task status, and all dates
                    msg.ClearTaskFlag()
                    
                    # Explicitly ensure FlagStatus is reset
                    msg.FlagStatus = 0 
                    
                    msg.Save()
                else:
                    print(f"Skipping non-email item (Class: {msg.Class})")
        except Exception as e:
            print(f"Error processing email: {e}")
    
    print("Email should have been unflagged")
    time.sleep(30)  # Wait a few seconds to ensure the email is sent before checking the inbox again

# -------------------------------
# Check and send follow-ups
# -------------------------------
def process_followups(df):
    now = datetime.now()

    for index, row in df.iterrows():
        if pd.isna(row["next_followup_due"]):
            continue

        if now >= row["next_followup_due"]:
            send_email(row["email"], row['subject_line'])

            # Schedule next follow-up again (repeat cycle)
            df.at[index, "next_followup_due"] = now + timedelta(days=FOLLOWUP_DAYS)

    return df

# -------------------------------
# Main loop
# -------------------------------
def main():
    print("Starting Outlook automation...")

    init_csv()

    while True:
        try:
            df = load_data()

            df = scan_flagged_emails(df)
            df = process_followups(df)

            save_data(df)

            print(f"Checked at {datetime.now()}")

        except Exception as e:
            print(f"Main loop error: {e}")

        time.sleep(CHECK_INTERVAL)

# -------------------------------
# Run
# -------------------------------
if __name__ == "__main__":
    main()