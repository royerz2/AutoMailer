import random
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
from tqdm import tqdm

e = None  # Don't change this.

print('Loading...')

joke_mode = False  # You can change this if you dont want jokes. (not recommended, lol)
journal = True  # Change this to True if you are mailing on behalf of JOURNAL Team. False if funding.
add_brochure = False  # True attaches the brochure to the e-mail. False does nothing.
auto_mode = False  # Don't change this.

response = input('Do you want to check the mails before sending? (recommended) [y]/n')  # Response enables auto mail
if response == 'n':
    auto_mode = True  # Convert response to boolean

try:
    if journal:  # Login credentials for JOURNAL mail
        print('Logging in as JOURNAL Team...')
        sender = ''
        password = ''
        server = 'smtp.gmail.com'
        port = 465
    else:  
        print('Logging in as FUNDING Team...')
        sender = ""
        password = ''
        server = 'smtp.gmail.com'
        port = 465

    server = smtplib.SMTP_SSL(server, port)  # Sets server.
    server.ehlo()  # Ping
    server.login(sender, password)  # Login with the specified login information.
    server.ehlo()  # Ping

except Exception as e:
    print(e)
    print('Something went wrong contacting the server. Contact IT.')
    exit()

print('Logged in without issues.')

if joke_mode:
    jokes = {
        1: 'Lets hope they fund us well (pun).',
        2: 'Why couldn’t the nickel understand the dime? It wasn’t making any cents.',
        3: 'The duck will pay for your dinner and all you need to do is allow him to put it on his bill.',
        4: 'I was once a banker, but I lost interest.',
        5: 'What do you call an alligator that invests? Investygator',
        6: 'That mail was about iGEM? THATS A GEM!',
        7: 'What type of insect is worth money? A cent-ipede.',
        8: 'I hope that mail makes cents (pun)',
        9: 'I once invested in deers. I made a lot of bucks. (pun)',
        10: 'Which rabbit is rich enough to invest? Bucks Bunny.'
    }


def dictionary(name, correspondent, spesification, template):
    template_dict = {
        1: f"""\
Dear {name},
""",
    }
    return template_dict[template]


if add_brochure:
    brochure = '/Users/direduryan/PycharmProjects/iGEM-Mail/brochure.pdf'
    attach_file = open(brochure, 'rb')
    payload = MIMEBase('application', 'octate-stream')
    payload.set_payload(attach_file.read())
    encoders.encode_base64(payload)
    payload.add_header('Content-Disposition', 'attachment; filename=Brochure.pdf')


def construct_mail(index):
    template = df.template[index]  # Get the template of the indexed row.
    name = df.name[index]  # Get the name of the indexed row.
    correspondent = df.correspondent[index]  # Get the correspondent of the indexed row.
    spesification = df.spesification[index]  # Get the correspondent of the indexed row.
    body = dictionary(name, correspondent, spesification, template)  # Construct mail using the dict and index values.

    content = MIMEText(body, "plain")
    msg = MIMEMultipart()
    msg['To'] = df.mail[index]
    msg.attach(content)

    if journal:
        msg['Subject'] = 'Journal Initiative Participation'
        msg['From'] = 'iGEM Maastricht University'
    else:
        msg['Subject'] = 'iGEM Team Maastricht'
        msg['From'] = 'iGEM Maastricht University'
        msg.attach(payload)

    return body, msg


def send_mail(msg, index):
    try:
        server.sendmail(sender,
                        df.mail[index],
                        msg.as_string())
    except Exception as e:
        print(e)

        if joke_mode:
            print(jokes[random.randint(1, len(jokes))])


try:
    # Reads the .xls file of the test list and saves it as a dataframe.
    df = pd.read_excel('/Users/direduryan/PycharmProjects/iGEM-Mail/test.xls')

except Exception as e:
    print(e)
    print('Something went reading the spreadsheet. MAKE SURE THE SPREADSHEET IS SAVED AS .xls! If it is, contact IT.')
    exit()

print('Starting test sequence.')
for row in tqdm(range(df.shape[0])):

    text, to_be_sent = construct_mail(row)

    if auto_mode:
        send_mail(to_be_sent, row)

    else:
        print(text)
        confirmation = input(f'Do you confirm that this mail to be sent to {df.name[row]}? [y]/n ')

        if confirmation == 'y':
            send_mail(to_be_sent, row)

print('Test complete.')
test = input('Is the test successful? y/n')
test = "y"

if test == 'y':

    try:
        # Reads the .xls file of the mailing list and saves it as a dataframe.
        df = pd.read_excel('/Users/direduryan/PycharmProjects/iGEM-Mail/journal.xls')

    except Exception as e:
        print(e)
        print(
            'Something went reading the spreadsheet. MAKE SURE THE SPREADSHEET IS SAVED AS .xls! If it is, contact IT.')
        exit()

    print(df.head)

    input('Press enter to proceed...')

    for row in tqdm(range(df.shape[0])):

        text, to_be_sent = construct_mail(row)

        if auto_mode:
            print(f'Sending to {df.name[row]}')
            send_mail(to_be_sent, row)

        else:
            print(text)
            confirmation = input(f'Do you confirm that this mail to be sent to {df.name[row]}? [y]/n ')

            if confirmation == 'y':
                send_mail(to_be_sent, row)

print('All mails sent, quitting server.')
