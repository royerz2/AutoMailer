import mailer as rm

mailText = """
Hello some person,

This is a test email to test the capabilities of the mailer lib.
Please find attached an attachment. I don't know what rn. I'm
writing this beforehand you see.

Regards,
Utku 
"""
emails = []
attachment_files = []

mail = rm.mailer("smtp.gmail.com", password ,"utkuturker@gmail.com")
mail.connectToServer()
mail.constructMail(mailText, mailFrom="Utku Turker", subject="Not very important", attachDir=attachment_files)
for address in emails:
    mail.sendMail(address)

