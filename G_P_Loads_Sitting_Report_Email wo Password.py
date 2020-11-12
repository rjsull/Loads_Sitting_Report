# Python code to illustrate Sending mail with attachments 
# from your Gmail account  
  
# libraries to be imported 
import os
import smtplib 
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText
#from email.mime.image import MIMEImage
from email.mime.base import MIMEBase 
from email import encoders 

fromaddr = 'x'
toaddr = ['x','x']

# instance of MIMEMultipart 
msg = MIMEMultipart() 
  
# storing the senders email address   
msg['From'] = fromaddr 
  
# storing the receivers email address  
msg['To'] = ", ".join(toaddr) 
  
# storing the subject  
msg['Subject'] = "G&P Loads Sitting Report"
  
# string to store the body of the mail, includes a hyperlink 
body = """<pre> <font face="calibri" color="black" size="3">
Attached is the G&P Loads Sitting Report.

Thanks,
--
</font></pre>"""

# attach the body with the msg instance 
msg.attach(MIMEText(body, 'html')) 

#set working directory 
os.chdir('C:\\Users\\sullivanry\\Documents\\Loads-Sitting-Report')

# open the file to be sent, has to be in working directory, change file to invalid if you don't need attachments 
filename = "G_P_Loads_Sitting_Report.xlsx"
attachment = open(filename, "rb") 

# instance of MIMEBase and named as p 
p = MIMEBase('application', 'octet-stream') 

# To change the payload into encoded form 
p.set_payload((attachment).read()) 

# encode into base64 
encoders.encode_base64(p) 
p.add_header('Content-Disposition', "attachment; filename= %s" % filename) 
# attach the instance 'p' to instance 'msg' 
msg.attach(p) 
# creates SMTP session 
s = smtplib.SMTP('smtp.gmail.com', 587) 
  
# start TLS for security 
s.starttls() 

# Authentication, Enter Your Password Below
s.login(fromaddr, "x") 
  
# Converts the Multipart msg into a string 
text = msg.as_string() 
  
# sending the mail 
s.sendmail(fromaddr, toaddr, text) 
  
# terminating the session 
s.quit() 