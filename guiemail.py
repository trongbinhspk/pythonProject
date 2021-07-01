import smtplib
import mimetypes
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.message import Message
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
import xlrd

file_location="g:/abc.xlsx"
wb=xlrd.open_workbook(file_location)
sheet=wb.sheet_by_index(0)


emailfrom="trongbinhspk@gmail.com"
emailto=sheet.cell_value(0,1)
fileToSend="g:/HRM.docx"
username="trongbinhspk@gmail.com"
password="binhtrongtran"

msg=MIMEMultipart()
msg["From"]=emailfrom
msg["To"]=emailto
msg["Subject"]="Hopthưtựđộng"
msg.preamble="helpIcannotsendanattachmenttosavemylife"

html="""\
<html>
<head></head>
<body>
<p>Hi!<br>
Xinchào
</p>
</body>
</html>
"""

#RecordtheMIMEtypesofbothparts-text/plainandtext/html.
part2=MIMEText(html,'html')

#Attachpartsintomessagecontainer.
#AccordingtoRFC2046,thelastpartofamultipartmessage,inthiscase
#theHTMLmessage,isbestandpreferred.
msg.attach(part2)

ctype,encoding=mimetypes.guess_type(fileToSend)
if ctype is None or encoding is not None:
    ctype="application/octet-stream"

maintype,subtype=ctype.split("/",1)

if maintype=="text":
    fp=open(fileToSend)
    #Note:weshouldhandlecalculatingthecharset
    attachment=MIMEText(fp.read(),_subtype=subtype)
    fp.close()
elif maintype=="image":
    fp=open(fileToSend,"rb")
    attachment=MIMEImage(fp.read(),_subtype=subtype)
    fp.close()
elif maintype=="audio":
    fp=open(fileToSend,"rb")
    attachment=MIMEAudio(fp.read(),_subtype=subtype)
    fp.close()
else:
    fp=open(fileToSend,"rb")
    attachment=MIMEBase(maintype,subtype)
    attachment.set_payload(fp.read())
    fp.close()
encoders.encode_base64(attachment)
attachment.add_header("Content-Disposition","attachment",filename=fileToSend)
msg.attach(attachment)

server=smtplib.SMTP("smtp.gmail.com:587")
server.starttls()
server.login(username,password)
server.sendmail(emailfrom,emailto,msg.as_string())
server.quit()
