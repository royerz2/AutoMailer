import smtplib as smtp
import re
import os

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication


class LoginError(Exception): # Custom exception. Doesn't really have to exist since Except() works fine but do I look like I care?
    def __init__(self, user = None):
        self.user = user
        super().__init__(self.user)

    def __str__(self):
        return f"Couldn't log in with {self.user}. This might be due to a wrong username or password."


class mailer(): # This class is entirely responsible for mail related stuff. I think...
    def __init__(self,
                 smtpAddress,
                 smtpPassword = None,
                 smtpUsername = None,
                 smtpPort = 465):

        self.smtpServerPort = smtpPort
        self.smtpServerAddress = smtpAddress
        self.smtpUser = smtpUsername
        self.smtpPassword = smtpPassword
        self.smtpConnection = ""

        self.emailRegex = re.compile(r"([A-Za-z0-9]+[.-_])*[A-Za-z0-9]+@[A-Za-z0-9-]+(\.[A-Z|a-z]{2,})+") # Regex for email addresses

        self.mailBase = ""
        self.mailBase["Subject"] = ""
        self.mailBase["From"] = ""
        self.mailBase["To"] = ""

    def connectToServer(self): # Establish connection to the smtp server
        print("connectMethod")
        try:
            with smtp.SMTP_SSL(str(self.smtpServerAddress), int(self.smtpServerPort)) as self.smtpConnection:
                self.smtpConnection.login(self.smtpUser, self.smtpPassword)

        except smtp.SMTPAuthenticationError as e: raise LoginError(self.smtpUser)
        except Exception as e: raise Exception(f"Something went wrong?: {e}")

    def constructMail(self, mailTextContent = None, # Creates the structure of the email and populates it with the message
                      mailHtmlContent = None,
                      mailFrom = None,
                      subject = None,
                      attachDir = None):

        self.mailBase = MIMEMultipart("alternative") # Create a base for the email to be built on

        if mailTextContent is not None: self.mailBase.attach(MIMEText(mailTextContent, "plain", "utf-8")) # Embed text if provided
        if mailHtmlContent is not None: self.mailBase.attach(mailHtmlContent, "html", "utf-8") # Embed html if provided
        if subject is not None: self.mailBase["Subject"] = subject
        if mailFrom is not None:
            self.mailBase["From"] = mailFrom
        else:
            self.mailBase["From"] = self.smtpUser

        if attachDir is not None:
            for dirf in attachDir:
                try:  # Attach file(s) to the email if any provided
                    file = open(dirf,"rb")
                    mailAttachment = MIMEApplication(file.read())
                    mailAttachment.add_header("Content-Disposition", "attachment", (os.path.split(dirf)[1]))
                    self.mailBase.attach(mailAttachment)
                except Exception as e: # If file path doesn't exist, raise an error
                    pass

    def sendMail(self, # Send the email to the desired recipient
                 recipientAddr = None,
                 bcc = None):
        try:
            if recipientAddr is None: raise Exception("No address provided to send the email!")
            if not re.fullmatch(self.emailRegex, recipientAddr): raise Exception("The email address entered is not valid!")

            self.mailBase["To"] = recipientAddr

            self.smtpConnection.sendmail(self.mailBase["From"],
                                         self.mailBase["To"],
                                         self.mailBase)

        except smtp.SMTPServerDisconnected as e: raise Exception("A problem occurred while connecting to the server! Check if you are logged in!")
        except Exception as e: raise Exception(f"Something went wrong?: {e}")

    def terminateConnection(self):
        self.smtpConnection.quit()

if __name__ == "__main__": # Don't run this as a program ok?
    print("This program is not designed to be run!")
