import win32com.client
import pandas as pd
import os
import time
from datetime import datetime, timedelta

CSV_FILE = "email_tracking.csv"
CHECK_INTERVAL = 30  # seconds
FOLLOWUP_TIME_TYPE = "hours" # Change to "days", "hours", or "minutes" exactly
FOLLOWUP_TIME = 1 # Days, Hours, or Minutes until next follow-up if no response

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
            "sent_time_duration_type",
            "sent_time_duration_value",
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
# Helper: get all recipient emails from a message
# -------------------------------
def get_all_recipients(msg):
    """Returns a comma-separated string of all recipient email addresses."""
    all_recipients = []
    for i in range(1, msg.Recipients.Count + 1):
        r = msg.Recipients.Item(i)
        try:
            if r.AddressEntry.Type == "EX":
                all_recipients.append(r.AddressEntry.GetExchangeUser().PrimarySmtpAddress.lower())
            else:
                all_recipients.append(r.Address.lower())
        except:
            pass
    return ", ".join(all_recipients) if all_recipients else None


# -------------------------------
# Scan flagged emails today
# -------------------------------
def scan_flagged_emails(df):
    print("Scanning for flagged emails...")

    sent_items = get_outlook_sent_items()
    messages = sent_items.Items
    print(f"Total emails in sent items: {len(messages)}")

    messages.Sort("[ReceivedTime]", True)

    for msg in messages:
        try:
            received_time = msg.SentOn

            # CHANGED: collect all recipients instead of just the first one
            recipient_email = get_all_recipients(msg)

            if recipient_email is None:
                print("Skipping email with no recipients.")
                continue

            print(f"Scanning sent email to: {recipient_email}")

            if msg.FlagStatus == 2:
                print(f"Email to {recipient_email} with subject '{msg.Subject}' is flagged for follow-up.")

                if received_time.tzinfo is not None:
                    received_time = received_time.replace(tzinfo=None)

                now = datetime.now().replace(tzinfo=None)

                is_duplicate = (
                    (df["email"] == recipient_email) &
                    (df["subject_line"] == msg.Subject) &
                    (df["sent_date"] == received_time)
                ).any()

                if not is_duplicate:
                    print(f"Adding new entry for follow-up with {recipient_email}")
                    new_row = {
                        "email": recipient_email,
                        "subject_line": msg.Subject,
                        "sent_date": received_time,
                        "flagged_date": now,
                        "last_seen_date": now,
                        "sent_time_duration_type": FOLLOWUP_TIME_TYPE,
                        "sent_time_duration_value": FOLLOWUP_TIME,
                        "next_followup_due": now + timedelta(**{FOLLOWUP_TIME_TYPE: FOLLOWUP_TIME})
                    }

                    new_entry = pd.DataFrame([new_row])
                    to_concat = [df, new_entry]
                    valid_dfs = [d for d in to_concat if not d.empty]

                    if valid_dfs:
                        df = pd.concat(valid_dfs, ignore_index=True)

            else:
                if received_time.tzinfo is not None:
                    received_time = received_time.replace(tzinfo=None)

                is_logged = (
                    (df["email"] == recipient_email) &
                    (df["subject_line"] == msg.Subject) &
                    (df["sent_date"] == received_time)
                ).any()

                if is_logged:
                    df = df[~(
                        (df["email"] == recipient_email) &
                        (df["subject_line"] == msg.Subject) &
                        (df["sent_date"] == received_time)
                    )]
                    print(f"Removed {recipient_email} from tracking because it was unflagged.")

        except Exception as e:
            print(f"Error scanning email: {e}")

    return df


# -------------------------------
# Send follow-up email
# -------------------------------
def send_email(to_address, subject_line, sent_date, index, now, df):
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    # Replace all commas (",") with semicolons (";")
    formatted_emails = to_address.replace(",", ";")

    print("The email address is:", formatted_emails)
    # print("The email type is:", type(to_address))

    mail.To = formatted_emails  # Outlook handles comma-separated addresses natively
    mail.Subject = SUBJECT_TEXT
    mail.Body = BODY_TEXT

    mail.Send()
    print(f"Follow-up sent to {to_address}")

    sent_items = get_outlook_sent_items()
    messages = sent_items.Items
    messages.Sort("[ReceivedTime]", True)

    # CHANGED: split to_address so we can match any one of the recipients
    tracked_recipients = [e.strip() for e in to_address.split(",")]

    for msg in messages:
        try:
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

            received_time = msg.SentOn
            if received_time.tzinfo is not None:
                received_time = received_time.replace(tzinfo=None)

            # CHANGED: check if recipient is in the tracked list, not exact string match
            if recipient_email in tracked_recipients and msg.Subject == subject_line and sent_date == received_time:
                print(f"Found email to {recipient_email} with subject '{subject_line}' — unflagging.")
                # Update the next follow up based on the FOLLOWUP_TIME_TYPE and FOLLOWUP_TIME that was set when the email was flaggsed
                df.at[index, "next_followup_due"] = now + timedelta(**{df.at[index, "sent_time_duration_type"]: int(df.at[index, "sent_time_duration_value"])})
                print(f"For {recipient_email} setting next follow-up due to {df.at[index, 'next_followup_due']} based on sent_time_duration_type and sent_time_duration_value.")
                # Store in the CSV for the specific row
                df.to_csv("email_tracking.csv", index=False)
        except Exception as e:
            print(f"Error unflagging email: {e}")


# -------------------------------
# Check and send follow-ups
# -------------------------------
def process_followups(df):
    now = datetime.now()

    for index, row in df.iterrows():
        if pd.isna(row["next_followup_due"]):
            continue

        if now >= row["next_followup_due"]:
            # Save before sending to prevent duplicate emails on crash
            df.at[index, "next_followup_due"] = now + timedelta(**{FOLLOWUP_TIME_TYPE: FOLLOWUP_TIME})
            save_data(df)
            send_email(row["email"], row['subject_line'], row['sent_date'], index, now, df)

        inbox_messages = get_outlook_inbox()
        sent_items = get_outlook_sent_items()

        inbox_messages = inbox_messages.Items
        sent_messages = sent_items.Items

        inbox_messages.Sort("[ReceivedTime]", True)
        sent_messages.Sort("[ReceivedTime]", True)

        # CHANGED: split the stored comma-separated recipients into a list
        tracked_recipients = [e.strip() for e in row["email"].split(",")]

        for msg_inbox in inbox_messages:
            try:
                if msg_inbox.Class != 43:
                    continue

                current_sender_obj = getattr(msg_inbox, "Sender", None)
                if not current_sender_obj:
                    continue

                if msg_inbox.SenderEmailType == "EX":
                    sender = msg_inbox.Sender.GetExchangeUser().PrimarySmtpAddress.lower()
                else:
                    sender = msg_inbox.SenderEmailAddress.lower()

                print("Email in inbox:", sender)

                # CHANGED: check if sender is in the tracked recipients list
                if sender in tracked_recipients and row['subject_line'] in msg_inbox.Subject and now <= row["next_followup_due"]:
                    print(f"Reply detected from {sender} — stopping follow-up cycle.")

                    for msg_sent in sent_messages:
                        try:
                            recipient_email = None
                            if msg_sent.Recipients.Count > 0:
                                recipient = msg_sent.Recipients.Item(1)
                                try:
                                    if recipient.AddressEntry.Type == "EX":
                                        recipient_email = recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress.lower()
                                    else:
                                        recipient_email = recipient.Address.lower()
                                except:
                                    recipient_email = recipient.Address.lower()
                            else:
                                continue

                            received_time = msg_sent.SentOn
                            if received_time.tzinfo is not None:
                                received_time = received_time.replace(tzinfo=None)

                            # CHANGED: check if first recipient is in the tracked list
                            if recipient_email in tracked_recipients and row['subject_line'] == msg_sent.Subject and row['sent_date'] == received_time:
                                print(f"Unflagging original email to {recipient_email}...")
                                if msg_sent.Class == 43:
                                    msg_sent.ClearTaskFlag()
                                    msg_sent.FlagStatus = 0
                                    msg_sent.Save()

                        except Exception as e:
                            print(f"Error processing sent email: {e}")

            except Exception as e:
                print(f"Error processing inbox email: {e}")

    return df


# -------------------------------
# Main loop
# -------------------------------
def main():
    print("Starting Outlook automation...")
    
    # Validate configuration
    if FOLLOWUP_TIME_TYPE not in ["days", "hours", "minutes"]:
        raise ValueError(f"Invalid FOLLOWUP_TIME_TYPE: '{FOLLOWUP_TIME_TYPE}'. Must be 'days', 'hours', or 'minutes'.")

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