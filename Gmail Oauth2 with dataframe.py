#!/usr/bin/env python
# coding: utf-8

# In[9]:


import base64
import imaplib
import json
import smtplib
import urllib.parse
import urllib.request
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
import lxml.html
import pandas as pd 
from io import BytesIO
import io
from email.mime.nonmultipart import MIMENonMultipart
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.charset import Charset, BASE64
from email.mime.nonmultipart import MIMENonMultipart
from email import charset, encoders
from h3 import h3

GOOGLE_ACCOUNTS_BASE_URL = 'https://accounts.google.com'
REDIRECT_URI = 'urn:ietf:wg:oauth:2.0:oob'

## get your clientid and secret key- https://www.google.com/search?q=get+google+client+id+and+secret+key
## reference video - https://www.youtube.com/watch?v=9erstkJAuWI
GOOGLE_CLIENT_ID = 'ajibtfqbt0p2ujbtv47liq7mua.apps.googleusercontent.com'. # set your client id
GOOGLE_CLIENT_SECRET = '5Y72dqkGS42'  # set your client secret key

## set your refresh token value here , initially you should pass None as token to generate token.
GOOGLE_REFRESH_TOKEN= None
# GOOGLE_REFRESH_TOKEN = '1//0gNHDgQyUOY7-L9IrDeRA97Vbz-fRmNFXG4PhdFjEHVwCgYIARAAGBASNwF-_bQ4BT1jUPdKHOinagDZLO0CEN4ngzl8ikMnoal8'
print("imports done")


#reading multiple dataframes (here I am reading single set multiple times ) you can change as per your need.
final_booking1 = pd.read_excel('IBM_sample_employee_dataset.xlsx', sheet_name="Sheet1")
final_booking2 = pd.read_excel('IBM_sample_employee_dataset.xlsx', sheet_name="Sheet1")
final_booking3 = pd.read_excel('IBM_sample_employee_dataset.xlsx',sheet_name="Sheet1")


def command_to_url(command):
    return '%s/%s' % (GOOGLE_ACCOUNTS_BASE_URL, command)


def url_escape(text):
    return urllib.parse.quote(text, safe='~-._')


def url_unescape(text):
    return urllib.parse.unquote(text)


def url_format_params(params):
    param_fragments = []
    for param in sorted(params.items(), key=lambda x: x[0]):
        param_fragments.append('%s=%s' % (param[0], url_escape(param[1])))
    return '&'.join(param_fragments)


def generate_permission_url(client_id, scope='https://mail.google.com/'):
    params = {}
    params['client_id'] = client_id
    params['redirect_uri'] = REDIRECT_URI
    params['scope'] = scope
    params['response_type'] = 'code'
    return '%s?%s' % (command_to_url('o/oauth2/auth'), url_format_params(params))


def call_authorize_tokens(client_id, client_secret, authorization_code):
    params = {}
    params['client_id'] = client_id
    params['client_secret'] = client_secret
    params['code'] = authorization_code
    params['redirect_uri'] = REDIRECT_URI
    params['grant_type'] = 'authorization_code'
    request_url = command_to_url('o/oauth2/token')
    response = urllib.request.urlopen(request_url, urllib.parse.urlencode(params).encode('UTF-8')).read().decode('UTF-8')
    return json.loads(response)


def call_refresh_token(client_id, client_secret, refresh_token):
    params = {}
    params['client_id'] = client_id
    params['client_secret'] = client_secret
    params['refresh_token'] = refresh_token
    params['grant_type'] = 'refresh_token'
    request_url = command_to_url('o/oauth2/token')
    response = urllib.request.urlopen(request_url, urllib.parse.urlencode(params).encode('UTF-8')).read().decode('UTF-8')
    return json.loads(response)


def generate_oauth2_string(username, access_token, as_base64=False):
    auth_string = 'user=%s\1auth=Bearer %s\1\1' % (username, access_token)
    if as_base64:
        auth_string = base64.b64encode(auth_string.encode('ascii')).decode('ascii')
    return auth_string


def test_imap(user, auth_string):
    imap_conn = imaplib.IMAP4_SSL('imap.gmail.com')
    imap_conn.debug = 4
    imap_conn.authenticate('XOAUTH2', lambda x: auth_string)
    imap_conn.select('INBOX')


def test_smpt(user, base64_auth_string):
    smtp_conn = smtplib.SMTP('smtp.gmail.com', 587)
    smtp_conn.set_debuglevel(True)
    smtp_conn.ehlo('test')
    smtp_conn.starttls()
    smtp_conn.docmd('AUTH', 'XOAUTH2 ' + base64_auth_string)


def get_authorization(google_client_id, google_client_secret):
    scope = "https://mail.google.com/"
    print('Navigate to the following URL to auth:', generate_permission_url(google_client_id, scope))
    authorization_code = input('Enter verification code: ')
    response = call_authorize_tokens(google_client_id, google_client_secret, authorization_code)
#     print("refresh_token is ",response['refresh_token'],"access_token is ",response['access_token'],"expires in is ",response['expires_in'])
    return response['refresh_token'], response['access_token'], response['expires_in']


def refresh_authorization(google_client_id, google_client_secret, refresh_token):
    response = call_refresh_token(google_client_id, google_client_secret, refresh_token)
    return response['access_token'], response['expires_in']

def export_csv(df):
    with io.StringIO() as buffer:
        df.to_csv(buffer,index=False)
        return buffer.getvalue()


def export_excel(dfs):
    with io.BytesIO() as buffer:
        writer = pd.ExcelWriter(buffer)
        for df in dfs:
            df[0].to_excel(writer, index=False, sheet_name=df[1]) 
        writer.save()
    
        return buffer.getvalue()

def export_file(file):
    return open(report1['filename'], "rb").read()


report1 = {

    'content': export_excel([(final_booking1, 'Sheet1'), 
                             (final_booking2, 'Sheet2'),
                             (final_booking3, 'Sheet3')]), #     'export_csv(data)/export_excel_excel
    'filename': 'OutputFile.xlsx' #file_name_here 
}


def send_mail(fromaddr, toaddr, subject, message):
    access_token, expires_in = refresh_authorization(GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET, GOOGLE_REFRESH_TOKEN)
    auth_string = generate_oauth2_string(fromaddr, access_token, as_base64=True)
    
    msg = MIMEMultipart('related')

    msg['Subject'] = subject #write your subject line here
    msg['From'] = fromaddr
    msg['To'] = ",".join(toaddr)
#     msg['To'] = recipients
    msg.preamble = 'This is a multi-part message in MIME format.'
    msg_alternative = MIMEMultipart('alternative')
    msg.attach(msg_alternative)
    part_text = MIMEText(lxml.html.fromstring(message).text_content().encode('utf-8'), 'plain', _charset='utf-8')
    part_html = MIMEText(message.encode('utf-8'), 'html', _charset='utf-8')
    msg_alternative.attach(part_text)
    msg_alternative.attach(part_html)
    
    # Create the attachment of the message in text/csv.
#     attachment = MIMENonMultipart('text', 'csv', charset='utf-8')
#     attachment.add_header('Content-Disposition', 'attachment', filename=report1['filename'])
#     cs = Charset('utf-8')
#     cs.body_encoding = BASE64
#     attachment.set_payload(report1['content'].encode('utf-8'), charset=cs)
#     msg_alternative.attach(attachment)
    #attachment_code ends
    
    # Create xls attachment, this is your attachment.
    attachment = MIMEBase('application', "octet-stream")
    attachment.set_payload(report1['content'])
    encoders.encode_base64(attachment)
    attachment.add_header('Content-Disposition', 'attachment; filename="{}"'.format(report1['filename']))
    msg.attach(attachment)
    
    
    
    server = smtplib.SMTP('smtp.gmail.com:587')
    server.ehlo(GOOGLE_CLIENT_ID)
    server.starttls()
    server.docmd('AUTH', 'XOAUTH2 ' + auth_string)
    server.sendmail(fromaddr, toaddr, msg.as_string())
    server.quit()
    



if __name__ == '__main__':
    if GOOGLE_REFRESH_TOKEN is None:
        print('No refresh token found, obtaining one \n')
        refresh_token, access_token, expires_in = get_authorization(GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET)
        print('Set the following as your GOOGLE_REFRESH_TOKEN:\n', refresh_token)
        pass
    else:
        send_mail('sender@email.com',['receiver1@gmail.com','receiver1@gmail.com'],
                  'This is your subject line',
                  '<bThis is your mail body</b><br><br>' +
                  'So happy to hear from you!')
        print("Mail Sent")
