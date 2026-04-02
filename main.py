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
# Get Outlook Sent Items
# -------------------------------
def get_outlook_sent_items():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    sent_items = outlook.GetDefaultFolder(5)  # 5 = Sent Items
    return sent_items


# -------------------------------
# Scan flagged emails today
# -------------------------------
def scan_flagged_emails(df):
    print("Scanning for flagged emails...")
    
    # Get the Outlook sent items folder
    sent_items = get_outlook_sent_items()
    
    # Get all messages (emails) from the inbox
    messages = sent_items.Items
    print(f"Total emails in inbox: {len(messages)}")
    
    # Sort emails by received time in descending order (newest first)
    messages.Sort("[ReceivedTime]", True)

    # Loop through each email in the Sent Items folder
    for msg in messages:
        try:
            # 1. Use SentOn instead of ReceivedTime for Sent Items
            received_time = msg.SentOn

            # Get the recipient's email address (not sender, since this is Sent Items)
            sender = None
            if msg.Recipients.Count > 0:
                recipient = msg.Recipients.Item(1)
                try:
                    # For Exchange users, get the PrimarySmtpAddress
                    if recipient.AddressEntry.Type == "EX":
                        sender = recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress.lower()
                    else:
                        sender = recipient.Address.lower()
                except:
                    sender = recipient.Address.lower()
            
            if sender is None:
                print("Skipping email with no recipients.")
                continue
            
            print(f"Scanning sent email to: {sender}")

            # Check if this email has been flagged for follow-up
            if msg.FlagStatus == 2:
                print(f"Email to {sender} with subject '{msg.Subject}' is flagged for follow-up.")
                # print(f"Email to {sender} with subject '{msg.Subject}' status of flag is:", msg.flagstatus)
                # Force 'received_time' (SentOn) to be naive
                if received_time.tzinfo is not None:
                    received_time = received_time.replace(tzinfo=None)
                
                now = datetime.now().replace(tzinfo=None)

                # Check if this EXACT subject to this EXACT person is already logged
                is_duplicate = ((df["email"] == sender) & (df["subject_line"] == msg.Subject) & (df["sent_date"] == received_time)).any()

                if not is_duplicate:
                    print(f"Adding new entry for follow-up with {sender}")
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
                # If they are unflagged, remove from tracking
                # Use the 'sender' variable which now holds the recipient's address
                # Force the email's received_time to be naive to match the CSV column
                if received_time.tzinfo is not None:
                    received_time = received_time.replace(tzinfo=None)

                # Now calculate is_logged
                is_logged = ((df["email"] == sender) & 
                            (df["subject_line"] == msg.Subject) & 
                            (df["sent_date"] == received_time)).any()

                if is_logged:
                    # Remove only that specific email thread
                    df = df[~((df["email"] == sender) & 
                            (df["subject_line"] == msg.Subject) & 
                            (df["sent_date"] == received_time))]
                    print(f"Removed {sender} from tracking because it was unflagged.")
                # else:
                #     print("Received_time on file is:\n", df["sent_date"])
                #     print("Received_time of email is:", received_time)
                    # print(f"Email to {sender} with subject '{msg.Subject}', recived_time on file is:", df["sent_date"], "and:", received_time)

        except Exception as e:
            print(f"Error processing email: {e}")

    # Return the updated dataframe with all flagged emails
    return df


# -------------------------------
# Send follow-up email
# -------------------------------
def send_email(to_address, subject_line, sent_date):
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    mail.To = to_address
    mail.Subject = SUBJECT_TEXT
    mail.Body = BODY_TEXT

    mail.Send()
    print(f"Follow-up sent to {to_address}")

    # print("Email has been sent, going to unflagged the email in the inbox...")
    # time.sleep(30)  # Wait a few seconds to ensure the email is sent before checking the inbox again

    # Get the Outlook sent items folder
    sent_items = get_outlook_sent_items()
    
    # Get all messages (emails) from the inbox
    messages = sent_items.Items
    
    # Sort emails by received time in descending order (newest first)
    messages.Sort("[ReceivedTime]", True)

    # Loop through each email in the inbox
    for msg in messages:
        try:
            # Get the recipient's email address from Sent Items
            recipient_email = None
            if msg.Recipients.Count > 0:
                recipient = msg.Recipients.Item(1)
                try:
                    if recipient.AddressEntry.Type == "EX":
                        recipient_email = recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress.lower()
                    else:
                        recipient_email = recipient.Address.lower()
                except:
                    recipient_email = recipient.Address.lower()
            else:
                continue
            
            # Get the email's sent time (SentOn) and make it naive for comparison
            received_time = msg.SentOn
            if received_time.tzinfo is not None:
                received_time = received_time.replace(tzinfo=None)
            if recipient_email == to_address and msg.Subject == subject_line and sent_date == received_time:
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
    
    # print("Email should have been unflagged")
    # time.sleep(30)  # Wait a few seconds to ensure the email is sent before checking the inbox again


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