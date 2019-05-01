from O365 import Connection, MSGraphProtocol, Message, MailBox, Account
import re
import datetime
import json

with open('email.secrets.json', 'r') as f:
        secrets = json.loads(f.read()) 

CLIENT_ID = secrets['email_client_id']
CLIENT_SECRET = secrets['email_client_secret']
    
prot = MSGraphProtocol()
con = Connection((CLIENT_ID, CLIENT_SECRET))

def auth_user():
        credentials = (CLIENT_ID, CLIENT_SECRET)

        scopes = ['basic', 'message_all']
        account = Account(credentials)
        if not account.is_authenticated:
                account.authenticate(scopes=scopes)    


def get_emails():
        auth_user()
        mailbox = MailBox(con=con, protocol=prot)
        inbox = mailbox.inbox_folder()
        for message in inbox.get_messages():
                if 'bokning av thai' in message.subject.lower():
                        res = re.findall(r'(?<=dag )(.*)(?=<\/h2)', message.body, re.MULTILINE)[0].split(' ')
                        day = res[0]
                        month = res[1]
                        time = res[3]
                        print(str(datetime.datetime.now()))
                        print("Day: " + day + " Month: " + month + " time: " + time)


get_emails()