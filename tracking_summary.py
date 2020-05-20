import win32com.client as win32
import time
import pickle
import xml.etree.ElementTree as ET
from zeep import Client, Settings
from zeep.exceptions import Fault, TransportError, XMLSyntaxError
from datetime import datetime
import re
import sys

# Load the tracking numbers
with open('tracking.pkl', 'rb') as f:
    tracking_nums = pickle.load(f)

# Load credentials
with open('config.pkl', 'rb') as f:
    config = pickle.load(f)

license_number = config['license_number']
user_id = config['user_id']
password = config['password']
email = config['email']


if tracking_nums == {}:
    summary = 'No pending inbound shpiments'
    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.To = email
    mail.Subject = 'RMA Tracking Report'
    mail.Body = 'No pending shipments'
    mail.Send()
    sys.exit(email(summary))
else:
    pass   


# Set Connection
settings = Settings(strict=False, xml_huge_tree=True)
client = Client('c:\\Users\\24866\\Documents\\python\\ups\\Track.wsdl', settings=settings)

# Set SOAP headers
headers = {
    'UPSSecurity': {
        'UsernameToken': {
            'Username': user_id,
            'Password': password
        },
        'ServiceAccessToken': {
            'AccessLicenseNumber': license_number
        }
    }
}

# Create request dictionary
requestDictionary = {"RequestOption": "15",
                     "SubVersion":"1707"
                    }
trackingOption = "02"
upsLocale = "en_US"


def track_package(tracking_num):
# Try operation
# Sleep 2 seconds inbetween queries
    try:
        response = client.service.ProcessTrack(_soapheaders=headers, Request=requestDictionary,
                                            InquiryNumber=tracking_num,TrackingOption=trackingOption,Locale=upsLocale)
        if response['Response']['ResponseStatus']['Description'] == 'Success':
            return response

    except Fault:
        response = 'error'
        return response


def small_package(rma, tracking, response):
    # Package status
    if response['Shipment'][0]['Package'][0]['Activity'][0]['Status']['Description'] == 'ORIGIN SCAN':
        return f'RMA {rma} with tracking number {tracking} has been scanned but not picked up'
    elif response['Shipment'][0]['Package'][0]['Activity'][0]['Status']['Description'] == 'IN TRANSIT':
        # Need to test in transit number
        pass
    elif response['Shipment'][0]['Package'][0]['Activity'][0]['Status']['Description'] == 'DELIVERED':
        month = response['Shipment'][0]['Package'][0]['Activity'][0]['Date'][4:6]
        day = response['Shipment'][0]['Package'][0]['Activity'][0]['Date'][-2:]
        year  = response['Shipment'][0]['Package'][0]['Activity'][0]['Date'][:4]
        return f'RMA {rma} with tracking number {tracking} was delivered on {month + "-" + day + "-" + year}'
  

def freight(rma, tracking, response):
    if response['Shipment'][0]['CurrentStatus']['Description'] == 'In Transit':
        type = response['Shipment'][0]['DeliveryDetail'][0]['Type']['Description']
        date = response['Shipment'][0]['DeliveryDetail'][0]['Date']
        return f'RMA {rma} with tracking number {tracking} has {type} date {datetime.strftime(datetime.strptime(date, "%Y%m%d"), "%a %m-%d-%Y")}'
    elif response['Shipment'][0]['CurrentStatus']['Description'] == 'Delivered':
        pass
    #elif 


summary = ''
for rma, tracking in tracking_nums.items():
    response = track_package(tracking)
    if response == 'error':
        summary += f'RMA {rma} with tracking number {tracking} is invalid. Please verify tracking'  + '\n'
    elif not response['Shipment'][0]['ShipmentType']:
        summary += f'RMA {rma} with tracking number {tracking} hasn\'t been picked up yet' + '\n'
    elif response['Shipment'][0]['ShipmentType']['Description'] == 'Small Package':
        status = small_package(rma, tracking, response)
        summary += status  + '\n'
        time.sleep(2)
    elif response['Shipment'][0]['ShipmentType']['Description'] == 'Freight':
        status = freight(rma, tracking, response)
        summary += status  + '\n'
        time.sleep(2)

# Email summary
outlook = win32.Dispatch('Outlook.Application')
mail = outlook.CreateItem(0)
mail.To = 'erik.ottley@wdc.com'
mail.Subject = 'RMA Tracking Report'
mail.Body = f"""
Below are the RMA's that are in transit or delivered to your site.

{summary}
Please confirm the deliveries.

Thank you
"""

mail.Send()

# Remove delivered items from dictionary
term = re.compile(r'(RMA\s)(\d+)\s')
with open('tracking.pkl', 'rb') as f:
    tracking_nums = pickle.load(f)
    
for line in summary.split('\n'):
    if 'delivered' in line:
        key = re.findall(term, line)[0][1]
        del tracking_nums[key]

with open('tracking.pkl', 'wb') as f:    
    pickle.dump(tracking_nums, f)
