import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from jinja2 import Template
from pprint import pprint
from time import sleep
import pandas as pd
from pprint import pprint


def getUsers(fileName):
    xl = pd.ExcelFile(fileName)
    df = xl.parse()
    df_dict = df.to_dict(orient='records')
    return df_dict



def sendMailToUsers(sender, instList, subject):
    # due to limitations on how many emails can be sent in an hour, 
    for institution in instList:

        # create email list from instList.
        emails = institution['Email'].split(',')
        # create you html template here.  Just place your html inbetween the """ """. Use an html online editor to get an idea of what
        # email will look like if necessary.
        body = Template('''
<html>
    <link rel="stylesheet" href="https://www.w3schools.com/w3css/4/w3.css">
    <head>
        <style>
            table {
                font-family: arial, sans-serif;
                border-collapse: collapse;
                width: 100%;
            }

            td, th {
                border: 1px solid #dddddd;
                text-align: left;
                padding: 8px;
            }

            tr:nth-child(even) {
                background-color: #dddddd;
            }

            tr:nth-child(odd) {
                background-color: #dddddd;
            }

            tr:first-child {
                background-color: #e8772a;
            }

            #div-one {
                background-color: #1f1f2e;
                color: white;
            }

            #div-two {
                background-color: white;
            }

            #div-three {
                background-color: #DAD8D8;
                text-align: center;
                height: 20px;
            }

            img {
                display: block;
                margin-left: auto;
                margin-right: auto;
                width: 60%;
            }
        </style>
    </head>
    <body>
        <div id="div-two">
            <p> Hello!</p>

            <p>See below for your weekly switch numbers. This shows last week's numbers vs the week prior.</p>

            <table>
                <tr>
                    <th>Week</th>
                    <th>Enrolled</th>
                    <th>Created</th>
                    <th>Submitted</th>
                </tr>
                    <tr>
                        <td>
                            <p>{{institution['Prev.Date']}}</p>
                        </td>
                        <td>
                            <p>{{institution['Prev.Enrolled']}}</p>
                        </td>
                        <td>
                            <p>{{institution['Prev.Created']}}</p>
                        </td>
                        <td>
                            <p>{{institution['Prev.Submitted']}}</p>
                        </td>
                    </tr>
                     <tr>
                        <td>
                            <p>{{institution['Cur.Date']}}</p>
                        </td>
                        <td>
                            <p>{{institution['Cur.Enrolled']}}</p>
                        </td>
                        <td>
                            <p>{{institution['Cur.Created']}}</p>
                        </td>
                        <td>
                            <p>{{institution['Cur.Submitted']}}</p>
                        </td>
                    </tr>
            </table>
            <p>Please feel free to contact your Account Manager for help at AM@clickswitch.com.</p>
            <p>TIP: Keep ClickSWITCH top of mind with your staff by doing these few steps:</p>

            <p><strong>1. Enroll Everyone</strong></p>
            <p><strong>2. Give the account holder their switch track code</strong></p>
            <p><strong>3. Go for the Direct Deposit!</strong></p>

            <p>Then...</p>
            <p>Hold your staff accountable to their progress. Feel free to forward these numbers each week to your branch managers/staff so that they will see how they are doing.</p>
            <p>Again, feel free to reach out to Account Management or Support at support@clickswitch.com for any help or retraining you would like.</p>

            <p>We are always here to help!</p>

            <p>Click on our logo to view upcoming training sessions!</p>
        </div>
        <div id="div-two">
            <a href="http://www.clickswitch.com/training">
                <img src="cid:image1">
            </a>
        </div>
    </body>
</html>
        ''')

        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = ", ".join(institution['Email'])
        msg['Subject'] = subject

        # body.render() This is where you will feed your template any variables used inside the template.
        #(ex. html = body.render(services=user['services']))  user['services'] is a list of dictionaries
        html = body.render(institution=institution)
        # Create a MIME text object using the rendered html template
        HTML_BODY = MIMEText(html, 'html')

        
        msg.attach(HTML_BODY)

        # open file containing image.
        fp = open('App/Resources/static/images/ClickSWITCH Logo - Color.jpg', 'rb')
        msgImage = MIMEImage(fp.read(), _subtype="jpeg")
        fp.close()

        # define the image's ID referenced in the html template above.
        msgImage.add_header('Content-ID', '<image1>')
        msg.attach(msgImage)

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login('samrelliner@gmail.com','Scherbatsky')
        text = msg.as_string()
        try:
            server.sendmail(sender, emails, text)
        except smtplib.SMTPRecipientsRefused:
            sleep(120)
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.sendmail(sender, emails,text)
        server.quit()

# getUsers('NewWeeklyNumbersTemplate.xlsx')

sendMailToUsers('srehlinger@clickswitch.com', getUsers('NewWeeklyNumbersTemplate.xlsx'), "Weekly Switch Numbers")
