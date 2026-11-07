# Insert your python code here

import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows


while True:    
    directory = single_input("Please enter the path to your folder.")
    directory = directory.replace('"','')
    if directory.endswith('\\') == True:
        pass
    else:
        directory = directory +'\\'
    try: 
        filenames = os.listdir(directory)
        if 'ProductData.xlsx' in filenames and 'MasterInvoice.pdf' in filenames and 'Order1.xlsx' in filenames and 'Order2.xlsx' in filenames and 'Order3.xlsx' in filenames and 'Order4.xlsx' in filenames:
            break
        else:
            show_message("Error", "The file you entered is not correct. Please try again.")
    except:
        show_message("Error", "The file you entered is not correct. Please try again.")

print(directory)


product_descriptions = {}
product_prices = {}


ProductData = 'ProductData.xlsx'
full_path = os.path.join(directory, ProductData)


df = pd.read_excel(full_path, sheet_name='Descriptions')
for index, row in df.iterrows():
    product_descriptions[row['Unit number']] = row['Description']


print(product_descriptions)


dfp = pd.read_excel(full_path, sheet_name = 'Prices')
for index, row in dfp.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']


print(product_prices)


order_data={}
template1 = openpyxl.load_workbook('Order1.xlsx', data_only= True)
def get_order_data(invoice_num,filename, directory):
    orderdict={}
    selected_sheet = template1['OrderForm']
    email = selected_sheet['C13'].value
    name = selected_sheet['C14'].value
    date = selected_sheet['C9'].value.strftime('%m/%d/%Y')
    Unit1 = selected_sheet['B24'].value
    Quantity1 = selected_sheet['C24'].value
    Unit2 = selected_sheet['B25'].value
    quantity2 = selected_sheet['C25'].value
    Unit3 = selected_sheet['B26'].value
    quantity3 = selected_sheet['C26'].value
    order_data['InvoiceNum']= invoice_num
    order_data['Date']=date
    if Unit3 is not None:
        order_data['Total'] = (product_prices.get(Unit1)*Quantity1) + ((product_prices.get(Unit2)*quantity2)) + (product_prices.get(unit3)*quantity3)
        order_data['Total']= f'${order_data['Total']:,.2f}'
        order_data['CustomerAddress1'] = selected_sheet['C16'].value
        order_data['CustomerName'] = name
        order_data['CustomerCompany'] = selected_sheet['C15'].value
        order_data['CustomerAddress2'] = f'{selected_sheet['C17'].value}, {selected_sheet['C18'].value}, {selected_sheet['C19'].value}'
        order_data['Price0']= product_prices.get(Unit1)
        order_data['Price0']= f'${order_data['Price0']:,.2f}'
        order_data['Quantity0'] = Quantity1
        order_data['Amount0'] = product_prices.get(Unit1) * Quantity1
        order_data['Amount0'] = f'${order_data['Amount0']:,.2f}'
        order_data['Description0'] = product_descriptions.get(Unit1)
        order_data['Price1']= product_prices.get(Unit2)
        order_data['Price1']= f'${order_data['Price1']:,.2f}'
        order_data['Quantity1'] = quantity2
        order_data['Amount1'] = product_prices.get(Unit2) * quantity2
        order_data['Amount1'] = f'${order_data['Amount1']:,.2f}'
        order_data['Description1'] = product_descriptions.get(Unit2)
        order_data['Price2']= product_prices.get(Unit3)
        order_data['Price2']= f'${order_data['Price2']:,.2f}'
        order_data['Quantity2'] = quantity3
        order_data['Amount2'] = product_prices.get(Unit3) * quantity3
        order_data['Amount2'] = f'${order_data['Amount2']:,.2f}'
        order_data['Description2'] = product_descriptions.get(Unit2)
    elif Unit2 is not None:
        order_data['Total'] = (product_prices.get(Unit1) * Quantity1) + ((product_prices.get(Unit2) * quantity2))
        order_data['Total'] = f'${order_data['Total']:,.2f}'
        order_data['CustomerAddress1'] = selected_sheet['C16'].value
        order_data['CustomerName'] = selected_sheet['C14'].value
        order_data['CustomerCompany'] = selected_sheet['C15'].value
        order_data['CustomerAddress2'] = f'{selected_sheet['C17'].value}, {selected_sheet['C18'].value}, {selected_sheet['C19'].value}'
        order_data['Price0'] = product_prices.get(Unit1)
        order_data['Price0'] = f'${order_data['Price0']:,.2f}'
        order_data['Quantity0'] = Quantity1
        order_data['Amount0'] = product_prices.get(Unit1) * Quantity1
        order_data['Amount0'] = f'${order_data['Amount0']:,.2f}'
        order_data['Description0'] = product_descriptions.get(Unit1)
        order_data['Price1'] = product_prices.get(Unit2)
        order_data['Price1'] = f'${order_data['Price1']:,.2f}'
        order_data['Quantity1'] = quantity2
        order_data['Amount1'] = product_prices.get(Unit2) * quantity2
        order_data['Amount1'] = f'${order_data['Amount1']:,.2f}'
        order_data['Description1'] = product_descriptions.get(Unit2)
    else:
        order_data['Total'] = (product_prices.get(Unit1) * Quantity1)
        order_data['Total'] = f'${order_data['Total']:,.2f}'
        order_data['CustomerAddress1'] = selected_sheet['C16'].value
        order_data['CustomerName'] = selected_sheet['C14'].value
        order_data['CustomerCompany'] = selected_sheet['C15'].value
        order_data['CustomerAddress2'] = f'{selected_sheet['C17'].value}, {selected_sheet['C18'].value}, {selected_sheet['C19'].value}'
        order_data['Price0'] = product_prices.get(Unit1)
        order_data['Price0'] = f'${order_data['Price0']:,.2f}'
        order_data['Quantity0'] = Quantity1
        order_data['Amount0'] = product_prices.get(Unit1) * Quantity1
        order_data['Amount0'] = f'${order_data['Amount0']:,.2f}'
        order_data['Description0'] = product_descriptions.get(Unit1)


    return  order_data, name, email, date

order_data, name, email, date = get_order_data(8515, 'Order1.xlsx', directory)

def create_invoice(invoice_num, order_data, directory):
    create_pdf('MasterInvoice.pdf', f'Invoice {invoice_num}.pdf', order_data)


create_invoice(8515, order_data, directory)

gmail_password_path = select_file('Please select the gmail password file')

gmail_password = open(gmail_password_path).read()
def send_email(invoice_num, name, email, date, directory):
    gmail_user = 'muirkira@gmail.com'



    recipients = email


    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'


    body = f"""Dear {name},
   
    Attached is the invoice dated {date}.


    Thank you for your business.
   
    Python Tech Co.
    10 Python Drive
    Pleasant Grove, UT 84062
    Phone (801) 555-2372
    Email: sales@pythontechco.com"""


    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)


    filename = f'Invoice {invoice_num}.pdf'
    path_to_file = os.path.join(directory, filename)


    path_to_file = os.getcwd()+f'\\Invoice {invoice_num}.pdf'


    message = add_attachment(path_to_file, message)
       


    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)


# Send the email


    mail_server.send_message(message)
    mail_server.quit()


send_email(8515,'Thomas Holt','muirkira@gmail.com','11/13/2023',directory)

filenames = os.listdir(directory)
print(filenames)




invoice_num = 8515


for file in filenames:
    if 'Order' in file and 'xlsx' in file:
        order_data, name, email, date = get_order_data(invoice_num, file, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, 'muirkira@gmail.com', date, directory)
        invoice_num = invoice_num + 1