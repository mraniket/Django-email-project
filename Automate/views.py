from django.shortcuts import HttpResponse
from django.shortcuts import render
import openpyxl
import smtplib, ssl
import pandas as pd
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from django.template.loader import render_to_string
from django.utils.html import strip_tags


message = MIMEMultipart()

# start the email login part
smtp = smtplib.SMTP("smtp.outlook.com")
smtp.starttls()
smtp.login("aashinde20@gmail.com", "Satara@123")


def index(request):


    if "GET" == request.method:
        return render(request, 'index.html', {})
    else:
        excel_file = request.FILES["excel_file"]
        email_lists = pd.read_excel(excel_file)
        custom_message = request.POST['custom_message']
        print(custom_message)
        names = email_lists["NAME"]
        emails = email_lists["EMAIL"]
        status = email_lists["STATUS"]

        print("login success")

        excel_data = list()

        for i in range(len(emails)):
            name = names[i]
            email = emails[i]
            stts = status[i]
            msg = custom_message
            
            if stts == "Overdue":
                
                context ={
                        "name":name,
                        "email":email,
                        "status":stts,
                        "custom_message":msg

                    }
                excel_data.append(name)
                html_content = render_to_string("email.html", context)
                text_content = strip_tags(html_content)
                part2 = MIMEText(html_content, 'html')
                message = MIMEMultipart("alternative")
                message["Subject"] = "General Notice"
                message["From"] = "aashinde20@gmail.com"
                message["To"] = email
                message["CC"] = "sanketj019@gmail.com"
                message.attach(part2)
                # smtp.sendmail("aashinde20@gmail.com", [email], message.as_string())
                print(i)
               

            else:
                print("No overdues")

        
        return render(request, 'email.html', context)
        # return HttpResponse("Email Sent successfully")
