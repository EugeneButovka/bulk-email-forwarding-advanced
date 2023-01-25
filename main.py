import datetime
import imaplib
import email
import smtplib
import time

# from types import NoneType
from syjson import SyJson

IMAP_HOST = "imap.mail.ru"
SMTP_HOST = "smtp.mail.ru"
SMTP_PORT = 587

USERNAME = "your@email"
PASSWORD = "your.password"
SEARCH_CRITERIA = "ALL"#"FLAGGED"

FROM_ADDRESS = "from@email"
TO_ADDRESS = "to@email"

FORWARD_TIME_DELAY = 5
EXCEPTION_TIME_DELAY = 60
VERBOSE = True

# Open IMAP connection
imap_client = imaplib.IMAP4_SSL(IMAP_HOST)
imap_client.login(USERNAME, PASSWORD)

parsed_message_ids_json = SyJson(
    "message_ids.json",  # Path of the json file
    create_file=True,  # If the file does not exists,
    # this will automatically create that file
)
parsed_message_ids_json.create("ids", [])

messages_id_list = []
for bucket in ["INBOX"]:  # , "INBOX"]: # "&BB4EQgQ,BEAEMAQyBDsENQQ9BD0ESwQ1-"
    # Fetch messages' ID list
    status, _ = imap_client.select(bucket,
                                   readonly=True)  # .select("INBOX", readonly=True) #.select("&BB4EQgQ,BEAEMAQyBDsENQQ9BD0ESwQ1-", readonly=True)
    # status, _ = imap_client.select("INBOX", readonly=True)
    if status != "OK":
        raise Exception("Could not select connect to INBOX.")

    status, data = imap_client.uid("search",
                                   SEARCH_CRITERIA)  # .search(None, SEARCH_CRITERIA) #.sort("DATE",None,SEARCH_CRITERIA) #
    if status != "OK":
        raise Exception("Could not search for emails.")

    current_messages_id_list = data[0].decode("utf-8").split(' ')
    print("adding ", len(current_messages_id_list), "from", bucket)
    messages_id_list.extend(data[0].decode("utf-8").split(' '))

messages_id_list.sort(reverse=False, key=int)
messages_id_list_dupes = set([x for x in messages_id_list if messages_id_list.count(x) > 1])
print("id dupes count", len(messages_id_list_dupes), ":", messages_id_list_dupes)

if VERBOSE:
    print("dirs: " + imap_client.list().__str__())
    print("{} messages were found. Forwarding will start immediately.".format(len(messages_id_list)))
    print("Messages ids: {}".format(messages_id_list))
    print()

# messages_parsed = []
for msg_index, msg_id in enumerate(messages_id_list):
    # if msg_index > 3:
    #     break
    print()
    print()
    print("parsing message", msg_id)
    if parsed_message_ids_json["ids"].__contains__(msg_id):
        print("message already parsed")
        continue
    status, msg_data = imap_client.uid("fetch", msg_id, '(RFC822)')  # .fetch(msg_id, '(RFC822)')
    if status != "OK":
        raise Exception("Could not fetch email with id {}".format(msg_id))

    for response_part in msg_data:
        if isinstance(response_part, tuple):
            msg = email.message_from_bytes(response_part[1])
            letter_date = datetime.datetime(*email.utils.parsedate(msg["Date"])[:-2])

            # Change FROM and TO header of the message
            # msg.replace_header("From", FROM_ADDRESS)
            # msg.replace_header("To", TO_ADDRESS)

            # print(email.header.decode_header(msg["TO"]).__str__())
            # print(email.header.decode_header(msg["FROM"]).__str__())
            # print(email.header.decode_header(msg["Subject"]).__str__())

            try:
                to_field = email.header.decode_header(msg["TO"])[1][0].decode().__str__()
            except:
                try:
                    to_field = email.header.decode_header(msg["TO"])[0][0].__str__()
                except:
                    to_field = "No address"

            try:
                from_field = email.header.decode_header(msg["FROM"])[1][0].decode().__str__()
            except:
                try:
                    from_field = email.header.decode_header(msg["FROM"])[0][0].__str__()
                except:
                    from_field = "No address"

            try:
                subject_field = email.header.decode_header(msg["Subject"])[0][0].decode().__str__()
            except:
                try:
                    subject_field = email.header.decode_header(msg["Subject"])[0][0].__str__()
                except:
                    subject_field = "No subj"

            subject_field += " | TO: " + to_field + ", FROM: " + from_field + ", dated " + letter_date.__str__()

            print(to_field)
            print(from_field)
            print(subject_field)

            # Change FROM and TO header of the message
            try:
                msg.replace_header("From", FROM_ADDRESS)
            except:
                msg.add_header("From", FROM_ADDRESS)

            try:
                msg.replace_header("To", TO_ADDRESS)
            except:
                msg.add_header("To", TO_ADDRESS)

            try:
                msg.replace_header("Subject", subject_field)
            except:
                msg.add_header("Subject", subject_field)

            try:
                # Open SMTP connection
                smtp_client = smtplib.SMTP(SMTP_HOST, SMTP_PORT)
                smtp_client.starttls()
                smtp_client.ehlo()
                smtp_client.login(USERNAME, PASSWORD)

                # Send message
                smtp_client.sendmail(FROM_ADDRESS, TO_ADDRESS, msg.as_bytes())
                parsed_message_ids_json["ids"].append(msg_id)
                if VERBOSE:
                    print("Message {} was sent. {} emails from {} emails were forwarded."
                          .format(msg_id, len(parsed_message_ids_json["ids"]), len(messages_id_list)))

                # Close SMTP connection
                smtp_client.close()

                # Time delay until next command
                time.sleep(FORWARD_TIME_DELAY)
            except smtplib.SMTPSenderRefused as exception:
                if VERBOSE:
                    print("Encountered an error! Error: {}".format(exception))
                    print("Messages sent until now:")
                    print(parsed_message_ids_json["ids"])
                    print("Time to take a break. Will start again in {} seconds.".format(EXCEPTION_TIME_DELAY))
                time.sleep(EXCEPTION_TIME_DELAY)
            except smtplib.SMTPServerDisconnected as exception:
                if VERBOSE:
                    print("Server disconnected: {}".format(exception))
            except smtplib.SMTPNotSupportedError as exception:
                if VERBOSE:
                    print("Connection failed: {}".format(exception))
                    print("Messages sent until now:")
                    print(parsed_message_ids_json["ids"])
                    print("Time to take a break. Will start again in {} seconds.".format(EXCEPTION_TIME_DELAY))
                time.sleep(EXCEPTION_TIME_DELAY)
            except smtplib.SMTPDataError:
                raise Exception("Daily user sending quota exceeded.")

if VERBOSE:
    print("Job done. Enjoy your day!")
#
# # Logout
imap_client.close()
imap_client.logout()
