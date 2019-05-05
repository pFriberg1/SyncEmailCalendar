from O365 import Connection, MSGraphProtocol, Message, MailBox, Account
import re
from datetime import datetime, timedelta
import json
import calendar_helper

with open('email.secrets.json', 'r') as f:
        secrets = json.loads(f.read())

CLIENT_ID = secrets['email_client_id']
CLIENT_SECRET = secrets['email_client_secret']
MONTH_TO_NUM_MAP = {
       'januari': 1,
       'februari': 2,
       'mars': 3,
       'april': 4,
       'maj': 5,
       'juni': 6,
       'juli': 7,
       'augusti': 8,
       'september': 9,
       'oktober': 10,
       'november': 11,
       'december': 12
}

UPDATE_FREQUENCY = 1440

prot = MSGraphProtocol()
con = Connection((CLIENT_ID, CLIENT_SECRET))


def auth_user():
        credentials = (CLIENT_ID, CLIENT_SECRET)

        scopes = ['basic', 'message_all']
        account = Account(credentials)
        if not account.is_authenticated:
                account.authenticate(scopes=scopes)


def string_to_date(regex_str):
        day = int(regex_str[0])
        month = regex_str[1]
        time = regex_str[3].split(':')
        year = datetime.now().year
        hour = int(time[0])
        minutes = int(time[1])

        return datetime(year, MONTH_TO_NUM_MAP[month], day, hour, minutes)


def sort_messages(inbox):
        msg_list = []
        for message in inbox.get_messages():
                msg_list.append(message)
        msg_list.sort(key=lambda m: m.received)
        return msg_list


def get_emails():
        auth_user()
        mailbox = MailBox(con=con, protocol=prot)
        inbox = mailbox.inbox_folder()
        sorted_messages = sort_messages(inbox)

        for msg in sorted_messages:
                if msg.subject.lower().startswith('bokning av thai'):
                        res = re.findall(r'(?<=dag )(.*)(?=<\/h2)', msg.body,
                                re.MULTILINE)[0].split(' ')
                        #Convert date-string from emil to datetime
                        date = string_to_date(res)

                        #remove timezone info from the datetime the email was received
                        time_rec = msg.received
                        time_rec = time_rec.replace(tzinfo=None)
                       
                        if time_rec > (datetime.now() - timedelta(minutes=UPDATE_FREQUENCY)):
                                print('Bokning: ' + msg.subject + " Rec: " + str(msg.received) + " datum: " + str(date))
                                calendar_helper.book_muay_thai(date)

                elif msg.subject.lower().startswith('avbokning av thai'):
                        res = re.findall(r'(?<=dag )(.*)(?=<\/h2)', msg.body,
                                re.MULTILINE)[0].split(' ')

                        #Convert date-string from emil to datetime
                        date = string_to_date(res)

                        #remove timezone info from the datetime the email was received
                        time_rec = msg.received
                        time_rec = time_rec.replace(tzinfo=None)
                       
                        if time_rec > (datetime.now() - timedelta(minutes=UPDATE_FREQUENCY)):
                                print('Avbokning: ' + msg.subject + " Rec: " + str(msg.received) + " datum: " + str(date))
                                calendar_helper.delete_event(date)


get_emails()
