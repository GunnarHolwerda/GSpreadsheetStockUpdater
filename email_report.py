import smtplib
import authentication

fromaddr = authentication.from_addr
toaddr = authentication.to_addr
msg = "LOOK AT ME"

# Send the message via our own SMTP server
s = smtplib.SMTP('smtp.gmail.com:587')
s.starttls()
s.login(authentication.user, authentication.password)
s.sendmail(fromaddr, toaddr, msg)
s.quit()
