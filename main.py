import win32com.client
import pandas as pd
import os
import time
import pytz
from datetime import datetime, timedelta, timezone

CSV_FILE = "email_tracking.csv"
CHECK_INTERVAL = 30  # seconds
FOLLOWUP_DAYS = 7

SUBJECT_TEXT = "Follow up on the analysis report"
BODY_TEXT = "Hi, just following up on my previous message. Let me know when you get a chance."

# -------------------------------
# Initialize CSV if not exists
# -------------------------------
def init_csv():
    if not os.path.exists(CSV_FILE):
        df = pd.DataFrame(columns=[
            "email",
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

# global new_row

# -------------------------------
# Scan flagged emails today
# -------------------------------
def scan_flagged_emails(df):
    # Print status message to console
    print("Scanning for flagged emails...")
    
    # Get the Outlook inbox folder
    inbox = get_outlook_inbox()
    
    # Get all messages (emails) from the inbox
    messages = inbox.Items
    print(f"Total emails in inbox: {len(messages)}")
    
    # Sort emails by received time in descending order (newest first)
    messages.Sort("[ReceivedTime]", True)

    # Get today's date to filter only today's emails
    today = datetime.now().date()
    print(f"Today's date: {today}")

    # Loop through each email in the inbox
    for msg in messages:
        # Print debug info for each message
        print("Messages today:", msg)
        
        try:
            # Get the time the email was received
            received_time = msg.ReceivedTime
            
            # If the email is from before today, stop checking (older emails)
            if received_time.date() != today:
                break  # Stop the loop since we only want today's emails

            # Get the sender's email address and convert to lowercase
            sender = msg.SenderEmailAddress.lower()
            print(f"Scanning email from: {sender}")  # Print who we're checking

            # Check if this email has been flagged for follow-up (FlagStatus = 2 means flagged)
            if msg.FlagStatus == 2:  # 2 = flagged
                # Get the sender again (already done above, but kept for clarity)
                sender = msg.SenderEmailAddress.lower()

                # Get the current time
                # now = datetime.now()
                # 1. Define your timezone clearly

                # eastern = pytz.timezone('US/Eastern')
                # now = datetime.now(eastern)

                # Option A: Make everything timezone-aware (UTC)
                now = datetime.now().replace(tzinfo=None)
                # 2. Force 'received_time' to be naive
                if received_time.tzinfo is not None:
                    received_time = received_time.replace(tzinfo=None)
                # received_time = received_time.astimezone(timezone.utc) if received_time.tzinfo else received_time.replace(tzinfo=timezone.utc)

                # Check if this sender already exists in our tracking CSV
                if sender not in df["email"].values:
                    # If they exist, update their last seen date and schedule next follow-up
                    # df.loc[df["email"] == sender, "last_seen_date"] = now
                    # df.loc[df["email"] == sender, "next_followup_due"] = now + timedelta(days=FOLLOWUP_DAYS)
                # else:
                    # If they don't exist, create a new row with their info
                    new_row = {
                        "email": sender,
                        "sent_date": received_time,
                        "flagged_date": now,
                        "last_seen_date": now,
                        "next_followup_due": now + timedelta(days=FOLLOWUP_DAYS)
                    }
                    # new_entry = pd.DataFrame([new_row]) # Let pandas infer the correct datetime dtypes
                    # # Ensure existing df columns are datetimes so they match new_entry
                    # df['sent_date'] = pd.to_datetime(df['sent_date']) 
                    # df = pd.concat([df, new_entry], ignore_index=True)

                    # 2. DO NOT use .astype(object). Let pandas keep native datetime dtypes.
                    new_entry = pd.DataFrame([new_row])  # Create a DataFrame for the new row

                    # 3. Ensure existing 'df' columns match the timezone before concat
                    # If df is empty, this won't hurt. If it has data, it prevents alignment errors.
                    # for col in ["sent_date", "flagged_date", "last_seen_date", "next_followup_due"]:
                    #     if col in df.columns:
                    #         df[col] = pd.to_datetime(df[col]).dt.tz_convert('US/Eastern')

                    # df = pd.concat([df, new_entry], ignore_index=True)
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
                    # If they exist, update their last seen date and schedule next follow-up
                    if msg.FlagStatus != 2:
                        # remove from tracking if unflagged
                        df = df[df["email"] != sender]
                        print(f"Removed {sender} from tracking because it was unflagged")

        # If any error occurs while processing an email, print the error and continue
        except Exception as e:
            print(f"Error processing email: {e}")
            # print(f"Type of new_row: {type(new_row)})")

    # Return the updated dataframe with all flagged emails
    return df

# -------------------------------
# Send follow-up email
# -------------------------------
def send_email(to_address):
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    mail.To = to_address
    mail.Subject = SUBJECT_TEXT
    mail.Body = BODY_TEXT

    mail.Send()
    print(f"Follow-up sent to {to_address}")

# -------------------------------
# Check and send follow-ups
# -------------------------------
def process_followups(df):
    now = datetime.now()

    for index, row in df.iterrows():
        if pd.isna(row["next_followup_due"]):
            continue

        if now >= row["next_followup_due"]:
            send_email(row["email"])

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