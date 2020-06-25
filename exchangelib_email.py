from exchangelib import Credentials, Configuration, Account, DELEGATE
from exchangelib import Message, Mailbox, FileAttachment

from cfg import creds as cfg  # load your credentials


def send_email(account, subject, body, recipients, attachments=None):
    """
    Send an email.

    Parameters
    ----------
    account : Account object
    subject : str
    body : str
    recipients : list of str
        Each str is and email adress
    attachments : list of tuples or None
        (filename, binary contents)

    Examples
    --------
    >>> send_email(account, 'Subject line', 'Hello!', ['info@example.com'])
    """
    to_recipients = []
    for recipient in recipients:
        to_recipients.append(Mailbox(email_address=recipient))
    # Create message
    m = Message(account=account,
                folder=account.sent,
                subject=subject,
                body=body,
                to_recipients=to_recipients)

    # attach files
    for attachment_name, attachment_content in attachments or []:
        file = FileAttachment(name=attachment_name, content=attachment_content)
        m.attach(file)
    m.send_and_save()


credentials = Credentials(username=cfg['user'],
                             password=cfg['password'])

config = Configuration(server=cfg['server'], credentials=credentials)
account = Account(primary_smtp_address=cfg['smtp_address'], config=config,
                  autodiscover=False, access_type=DELEGATE)

# Read attachment
attachments = []
with open('exchangelib_email.py', 'rb') as f:
    content = f.read()
attachments.append(('exchangelib_email.py', content))

# Send email
send_email(account, 'Sent With the Automation', 'Just to be sure it works',
           ['xyz@test.com'],
           attachments=attachments)
