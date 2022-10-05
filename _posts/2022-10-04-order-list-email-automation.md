---
title: "Order List Email Automation"
date: 2022-10-04
categories:
  - blog
tags:
  - Python
  - Automation
---

Our hospitality director loves reaching out to folks immediately after making a purchase to give them a personalized thank you, provide greater detail about their product purchase, and to form a deeper connection - gotta love our focus on relationship building! I asked about her process of gathering purchaser information and it went as follows:
    - Log into back end of our e-commerce system
    - Click into each order individually
    - Copy email address
    - Open work email and paste email address into new message

While it may not seem like much work, she was doing this every morning and that time can add up, which could have certainly been used elsewhere. 

Thus, I thought it was best that I automate this process for her. My approach was as follows:
* Create connection to e-commerce system
* Extract only the most necessary information (e.g. name, email, product purchased)
* Gather extracted data into a csv file
* Email csv file with information for the orders purchased the day previously to her email every morning at 8:00 am

To begin, I'll import a few libraries and create a connection with the e-commerce REST API.
```python
# Import packages
import mimetypes
from woocommerce import API
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email import encoders
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
import smtplib, ssl
from datetime import date
from datetime import timedelta

# Create connection with e-commerce API. Key & Secret gathered from created REST API created in e-commerce back end. 
wcapi = API(
    url = "https://websiteaddress.com",
    consumer_key = "XXXXXXXXXXXXX"
    consumer_secret = "***********")

# Request API Connection
request = wcapi.get(wcapi)
json = request.json()
```

Once I formed the necessary connection, I began pulling only the information that she utilized when contacting purachasers. I decided on the product id, the date the purchase was made, first name, last name, email address, and product name.
```python
line_items = [order["line_items"] for order in json if isinstance(order, dict) and "line_items" in order]
new_dict = (
    ([order["id"] for order in json if isinstance(order, dict) and "id" in order],
     [order['date_created'][:-9] for order in json if isinstance(order, dict) and 'date_created' in order],
     [order['billing']['first_name'] for order in json if isinstance(order, dict) and 'billing' in order],
     [order['billing']['last_name'] for order in json if isinstance(order, dict) and 'billing' in order],
     [order['billing']['email'] for order in json if isinstance(order, dict) and 'billing' in order],
     [x[0]['name'] for x in line_items])
    [0:])
```

Next, since she only reached out to those folks who made purchases the previous day, I needed to ensure I was only pulling data from the day before (not the current day). Thus, I created a dataframe and cleaned it up so that it was easy to understand.  
```python
yesterday = pd.to_datetime(date.today() - timedelta(days=1))
df1 = pd.DataFrame(new_dict)
df2 = df1.transpose()
df2.columns = ['Order Number', 'Purchase Date', 'First Name', 'Last Name', 'Email', 'Course Name']
df2.index = df2.index + 1
df2['Purchase Date'] = pd.to_datetime(df2['Purchase Date'])
mask = (df2['Purchase Date'] == yesterday)
df2 = df2.loc[mask]
df2.to_csv('Product_Orders.csv', sep=',')
```

Next, I utilized gmail to send the csv file to her email. 
```python
def mail():
    msg = MIMEMultipart()
    body_part = MIMEText('Hello, Attached is a list of the orders made yesterday.', 'plain')
    msg['Subject'] = 'Daily Product Order Email'
    msg['From'] = 'from-email@gmail.com'
    msg['To'] = 'to-email@gmail.com'
    context = ssl.create_default_context()
    # Add body to email
    msg.attach(body_part)
    # attach .csv to email
    ctype, encoding = mimetypes.guess_type('Product_Orders.csv')
    if ctype is None or encoding is not None:
        ctype = 'application/octet-stream'
    maintype, subtype = ctype.split('/', 1)
    if maintype == "text":
        fp = open('Product_Orders.csv')
        # Note: we should handle calculating the charset
        attachment = MIMEText(fp.read(), _subtype=subtype)
        fp.close()
    elif maintype == "image":
        fp = open('Product_Orders.csv', "rb")
        attachment = MIMEImage(fp.read(), _subtype=subtype)
        fp.close()
    elif maintype == "audio":
        fp = open('Product_Orders.csv', "rb")
        attachment = MIMEAudio(fp.read(), _subtype=subtype)
        fp.close()
    else:
        fp = open('Product_Orders.csv', "rb")
        attachment = MIMEBase(maintype, subtype)
        attachment.set_payload(fp.read())
        fp.close()
        encoders.encode_base64(attachment)
    attachment.add_header("Content-Disposition", "attachment", filename='Product_Orders.csv')
    msg.attach(attachment)

    with open('Product_Orders.csv','rb') as file:
    # Attach the file with filename to the email
        msg.attach(MIMEApplication(file.read(), Name='Orders'))
    # Create SMTP object
    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp_obj:
# Login to the server
        smtp_obj.login('from-email@gmail.com', '***********')
# Convert the message to a string and send it
        smtp_obj.sendmail(msg['From'], msg['To'], msg.as_string())
        smtp_obj.quit()
mail()
```
Lastly, as a Mac user, I utilized Automator, an app native to the system, to schedule the code to be run each more at 8:00 am. I found setting up the Automator to be realitvely intuitive, however, there was a challenge that I believe needs to be emphasized:
* <strong>The Automator will not run the code if your computer goes to sleep or standby mode.</strong> 

You can change these settings, but I was only able to keep my laptop on so long as it was connected to a charger...something to keep in mind if you are to develop something similar. 

Thanks for sticking around! I appreciate you reading through how I've made a slight adjustment in the work that we do so to make a bigger impact on the clients we connect with. 

