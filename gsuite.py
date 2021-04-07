'''
   Demonstration how to fetch emails from gmail service
   and users info, attachment if mail has any
   Who is sender, which domain sender belongs etc..
'''
import base64
import re

from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials

VERIFY_SSL = False


class GsuiteService():
    scopes = ['https://www.googleapis.com/auth/admin.directory.group', \
              'https://www.googleapis.com/auth/admin.directory.user', \
              'https://www.googleapis.com/auth/admin.directory.domain', \
              'https://www.googleapis.com/auth/gmail.readonly']

    def __init__(self, credentials_file, admin_account):
        self.credentials_file = credentials_file
        self.admin_account = admin_account

    def messages(self):
        return self.Messages(self)

    def folders(self):
        return self.Folders(self)

    def groups(self):
        return self.Groups(self)

    def users(self):
        return self.Users(self)

    def domains(self):
        return self.Domains(self)


    class Groups():

        def __init__(self, outer):
            self.outer = outer
            credentials = ServiceAccountCredentials.from_json_keyfile_name(self.outer.credentials_file, scopes=self.outer.scopes)
            delegated_credentials = credentials.create_delegated(self.outer.admin_account)
            self.service = build('admin', 'directory_v1', credentials=delegated_credentials)

        def list(self):
            '''
            :return:
            '''

            try:
                response = self.service.groups().list(customer='my_customer').execute()
                groups_info = response.get('groups', [])
                groups = []
                if groups_info:
                    for group in groups_info:
                        groups.append(group.get("email"))

                    while 'nextPageToken' in response:
                        page_token = response['nextPageToken']
                        group_list = self.get_next_groups(page_token)
                        groups = groups + group_list
                    return (True, groups)
            except Exception as e:
                return (False, str(e))

        def check_group(self, ldap_group_name):
            """
               Return True of false based on ldap_group_name present or not.
               :return:
            """

            try:
                response = self.service.groups().get(groupKey=ldap_group_name).execute()
                group_id = response.get('id')
                return (True, group_id)
            except Exception as e:
                return (False, 'Group not Exists')

        def get_users(self, id):

            try:
                response = self.service.members().list(groupKey=id).execute()
                members_info = response.get('members', [])
                members = []
                if members_info:
                    for member in members_info:
                        members.append(member.get("email"))
                    while 'nextPageToken' in response:
                        page_token = response['nextPageToken']
                        member_list = self.get_next_users(id, page_token)
                        members = members + member_list
                    return (True, members)
            except Exception as e:
                return (False, str(e))

        def get_next_groups(self, page_token):

            group_list = []
            try:
                response = self.service.groups().list(customer='my_customer', pageToken=page_token).execute()
                if response.get('groups'):
                    for group in response.get('groups'):
                        group_list.append(group.get('email'))
                    return group_list
                else:
                    return None
            except Exception as e:
                return None

        def get_next_users(self, page_token):
            members_list = []
            try:
                response = self.service.members().list(groupKey=id, pageToken=page_token).execute()
                if response.get("members"):
                    for member in response.get("members"):
                        members_list.append(member.get('email'))
                    return members_list
                else:
                    return None
            except Exception as e:
                return None

    class Users():

        def __init__(self, outer):
            self.outer = outer
            credentials = ServiceAccountCredentials.from_json_keyfile_name(self.outer.credentials_file, scopes=self.outer.scopes)
            delegated_credentials = credentials.create_delegated(self.outer.admin_account)
            self.service = build('admin', 'directory_v1', credentials=delegated_credentials)

        def list(self, max_user=None, parameters=None):
            '''
            :return:
            '''

            try:
                # declasred on top
                if max_user:
                    results = self.service.users().list(customer='my_customer', maxResults=max_user).execute()
                else:
                    results = self.service.users().list(customer='my_customer').execute()
                users_info = results.get('users', [])
                users = []
                if users_info:
                    for user in users_info:
                        users.append(user.get("primaryEmail"))
                    return (True, users)
            except Exception as e:
                return (False, str(e))

        def check_user(self, email_id):
            """
               Return True of false based on user present or not.
               :return:
            """

            try:
                self.service.users().get(userKey=email_id).execute()
                return (True, 'User Exists')
            except Exception as e:
                return (False, 'User Not Exist')

    class Messages():

        def __init__(self, outer):
            self.outer = outer

        def list(self, user_context, query_param='', id=None, next_msgs_link=None):
            messages = []
            credentials = ServiceAccountCredentials.from_json_keyfile_name(self.outer.credentials_file, scopes=self.outer.scopes)
            delegated_credentials = credentials.create_delegated(user_context)

            service = build('gmail', 'v1',  credentials=delegated_credentials)
            try:
                if next_msgs_link is not None:
                    response = service.users().messages().list(userId=user_context, q=query_param, labelIds=['INBOX'],
                                                 pageToken=next_msgs_link, maxResults=100).execute()
                else:
                    response = service.users().messages().list(userId='me', labelIds=['INBOX'], maxResults=100).execute()
                if response.get('nextPageToken'):
                    next_page_token = response['nextPageToken']
                else:
                    next_page_token = None
                if 'messages' in response:
                    messages.extend(response['messages'])
                messages_content = []
                for message in messages:
                    msg = service.users().messages().get(userId='me', id=message['id']).execute()
                    messages_content.append(msg)

                return (True, messages_content, next_page_token)
            except Exception as e:
                return (False, str(e), next_msgs_link)

        def count(self, user_context, parameters=None):
            """
            Return count of messages
            :param user_context:
            :param parameters:
            :return:
            """
            #logger.debug('entering count message')

            credentials = ServiceAccountCredentials.from_json_keyfile_name(self.outer.credentials_file, scopes=self.outer.scopes)
            delegated_credentials = credentials.create_delegated(user_context)
            service = build('gmail', 'v1', credentials=delegated_credentials)
            try:
                profile = service.users().getProfile(userId=user_context).execute()
                return profile.get('messagesTotal', None)
            except Exception as e:
                print(str(e))
                return None

        def getattachment(self, user_context, message_id, parameters=None):
            """
                Retrieve a single message by id
                :param user_context:
                :param message_id:
                :param parameters:
                :return:
                """
            credentials = ServiceAccountCredentials.from_json_keyfile_name(self.outer.credentials_file, scopes=self.outer.scopes)
            delegated_credentials = credentials.create_delegated(user_context)
            service = build('gmail', 'v1',  credentials=delegated_credentials)
            attachments = []
            msg = service.users().messages().get(userId=user_context, id=message_id).execute()
            if msg['payload'].get('parts'):
                for part in msg['payload']['parts']:
                    if part.get("parts"):
                        for sub_part in part["parts"]:
                            attachment = self.get_attachments(service,sub_part, message_id)
                            if attachment:
                                attachments.append(attachment)
                    else:
                        attachment = self.get_attachments(service,part, message_id)
                        if attachment:
                            attachments.append(attachment)
                return attachments

        def get_attachments(self, service, part, message_id):
            store_dir = "/tmp/"

            if part['filename']:
                if 'data' in part['body']:
                    data = part['body']['data']
                else:
                    att_id = part['body']['attachmentId']
                    att = service.users().messages().attachments().get(userId='me', messageId=message_id,
                                                                       id=att_id).execute()
                    data = att['data']
                file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
                path = store_dir + part['filename']

                with open(path, 'wb') as f:
                    f.write(file_data)
                return path


#in place of magnetic-tenure-247105-11621e1d9402.json
#json credentials file 
#admin mail id 
service = GsuiteService('magnetic-tenure-247105-11621e1d9402.json', 'ta-eng@bce-demo.net')
results = service.users().list()
#fetch all the users
users = results[1]
global n_msgs_link

def fetch_messages(user_email, next_msgs_link=None):
    """
    Function to fetch messages of a o365 user.
    :param: email_id - string
    :return: Tuple (boolean, list)
    list of dict (messages)
    """
    if next_msgs_link is None:
        (is_successful, messages, next_msgs_link) = service.messages().list(user_email, id='Inbox',
    else:
        (is_successful, messages, next_msgs_link) = service.messages().list(user_email, next_msgs_link=next_msgs_link)

    return (is_successful, messages, next_msgs_link)


#fetch messages of each user
#fetch all the info of each mail
#like attachment, from , to etc..
for user in users:
    try:
        global n_msgs_link
        n_msgs_link = None
        success, msgs, next_msgs_link = fetch_messages(user, next_msgs_link=n_msgs_link)
    except Exception as e:
        print(str(e))

    count = 0
    if msgs:
        count = len(msgs)
        for msg in msgs:
            for header in msg.get('payload', {}).get('headers', []):
                if header.get('name', None) == 'From':
                    regx = '<(.*)>'

                    result = re.search(regx, header['value'])
                    if result:
                        print(result.groups(1)[0])
                        from_ = result.groups(1)[0]
                    else:
                        from_ = header.get('value')
                if header.get('name', None) == 'Subject':
                    msg_subject = header.get('value', None)
            frm_lower = from_.lower()

            for part in msg.get('payload', {}).get('parts', []):
                if part.get('mimeType', None) == 'text/html':
                    body_data = part.get('body', {}).get('data', None)
                    if body_data:
                        body_data = body_data.replace("-", "+")  # decoding from Base64 to UTF-8
                        body_data = body_data.replace("_", "/")  # decoding from Base64 to UTF-8
                        body_d = base64.b64decode(bytes(body_data, 'UTF-8'))  # decoding from Base64 to UTF-8

            attachments = service.messages().getattachment(user, msg['id'])
            if attachments:
                try:
                    for attachment in attachments:
                except Exception as e:
                    print(str(e))
            else:
                print("--- no attachments")
    print(count)
    print(service.messages().count(user))



