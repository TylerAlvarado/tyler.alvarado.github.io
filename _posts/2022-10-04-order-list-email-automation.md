---
title: "Order List Email Automation"
date: 2022-10-04
categories:
  - blog
tags:
  - Python
  - Automation
---

Develop connection with e-commerce API
```python
# Set-up WooCommerce API connection
from woocommerce import API
wcapi = API(
    url = "https://websiteaddress.com",
    consumer_key = "XXXXXXXXXXXXX"
    consumer_secret = "***********")

# Request API Connection
request = wcapi.get(wcapi)
json = request.json()
```

Gather Records from e-commerce using API connection
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

Filter Orders that occurred yesterday
```python
yesterday = pd.to_datetime(date.today() - timedelta(days=1))
df = pd.DataFrame(new_dict)
df1 = df.transpose()
df1.columns = ['Order Number', 'Purchase Date', 'First Name', 'Last Name', 'Email', 'Course Name']
df1.index = df1.index + 1
df1['Purchase Date'] = pd.to_datetime(df1['Purchase Date'])
mask = (df1['Purchase Date'] == yesterday)
df1 = df1.loc[mask]
df1.to_csv('Product_Orders.csv', sep=',')
```

Create a multipart email
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
