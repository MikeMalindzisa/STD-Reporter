import os
import shutil
import subprocess
import pandas as pd
import xml.etree.ElementTree as ET
import xml.dom.minidom as minidom
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import configparser
import logging
import pystray
from pystray import MenuItem as item
from PIL import Image
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import QFileSystemWatcher
from PyQt5.QtWidgets import QApplication
import datetime
# Set up logging
logger = logging.getLogger(__name__)
logger.setLevel(logging.ERROR)

# Create a file handler
log_file = 'error.log'
file_handler = logging.FileHandler(log_file)

# Set the log message format
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
file_handler.setFormatter(formatter)

# Add the file handler to the logger
logger.addHandler(file_handler)

# Get the path to the INI file
ini_path = 'stdsettings.ini'

# Create a ConfigParser instance
config = configparser.ConfigParser()

# Check if the INI file exists
if os.path.exists(ini_path):
    # Read the INI file
    config.read(ini_path)
else:
    # Create default folder names
    input_folder_name = 'input'
    output_folder_name = 'output'
    completed_folder_name = 'completed'
    failed_folder_name = 'failed'

    # Create default folder paths based on the current working directory
    input_folder = input_folder_name
    output_folder = output_folder_name
    completed_folder = completed_folder_name
    failed_folder = failed_folder_name

    # Create the default directories if they don't exist
    os.makedirs(input_folder, exist_ok=True)
    os.makedirs(output_folder, exist_ok=True)
    os.makedirs(completed_folder, exist_ok=True)
    os.makedirs(failed_folder, exist_ok=True)

    # Set the folder paths in the ConfigParser instance
    config['Folders'] = {
        'InputFolder': input_folder,
        'OutputFolder': output_folder,
        'CompletedFolder': completed_folder,
        'FailedFolder': failed_folder
    }

    # Write the ConfigParser instance to the INI file
    with open(ini_path, 'w') as config_file:
        config.write(config_file)

# Read the folder paths from the INI file
input_folder = config.get('Folders', 'InputFolder')
output_folder = config.get('Folders', 'OutputFolder')
completed_folder = config.get('Folders', 'CompletedFolder')
failed_folder = config.get('Folders', 'FailedFolder')

# Create the output folder if it doesn't exist
os.makedirs(output_folder, exist_ok=True)

# Create the completed folder if it doesn't exist
os.makedirs(completed_folder, exist_ok=True)

# Get the process ID of the running program
pid = os.getpid()

# Save the process ID to a file
pid_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), '_STD_REPORT_.pid')
with open(pid_file, 'w') as f:
    f.write(str(pid))


class FileHandler(FileSystemEventHandler):
    def on_created(self, event):
        if event.is_directory:
            return
        elif event.src_path.endswith('.xlsx'):
            process_excel(event.src_path, False)


def process_excel(excel_path, existing):
    try:
        # Rest of the code...try:
        # Read the Excel file into a pandas DataFrame
        df = pd.read_excel(excel_path)
        prefix = "00320231"
        current_datetime = datetime.datetime.now()
        unique_ref = prefix + current_datetime.strftime("%S%H%M")

        # Pad the unique_ref ID to a length of 14 characters
        # unique_ref = unique_ref.ljust(14, "0")
        # Create the root element of the XML structure
        root = ET.Element('Report')
        root.attrib = {'xmlns': 'http://transactionreporting.org/reportSchema',
                       'Reference': unique_ref,
                        'Environment': 'P',
                         'Version': '1'}

        count = 0
        for index, row in df.iterrows():

            try:
                line_number_value = count
                if pd.isna(line_number_value):
                    line_number = "Null"  # Set a default value or handle the missing value case
                else:
                    line_number = str(line_number_value + 1)

                # Populate the Record tags based on the columns in the Excel file
                original_transaction = ET.SubElement(root, 'OriginalTransaction')
            

                value_date = row['VALUE DATE']
                if pd.isna(value_date):
                    value_date_str = "Null"  # Set a default value or handle the missing value case
                else:
                    if pd.isnull(value_date):
                        value_date_str = "Null"  # Set a default value or handle the missing value case
                    else:
                        value_date_str = value_date.strftime('%Y-%m-%d').strip()

                

                dob = row['DATE OF BIRTH']
                if pd.isna(dob):
                    dob_str = "Null"  # Set a default value or handle the missing value case
                else:
                    if pd.isnull(dob):
                        dob_str = "Null"  # Set a default value or handle the missing value case
                    else:
                        dob_str = dob.strftime('%Y-%m-%d').strip()

                total_value = row['TOTAL VALUE']
                if pd.isna(total_value):
                    total_value_str = "Null"  # Set a default value or handle the missing value case
                else:
                    total_value = round(total_value)
                    total_value_str = str(total_value).strip()
                    
                
                foreign_value = row['FOREIGN VALUE ']
                if pd.isna(foreign_value):
                    foreign_value_str = "Null"  # Set a default value or handle the missing value case
                else:
                    foreign_value = round(foreign_value)
                    foreign_value_str = str(foreign_value).strip()
                    

                
                domestic_value = row['DOMESTIC VALUE ']
                if pd.isna(domestic_value):
                    domestic_value_str = "Null"  # Set a default value or handle the missing value case
                else:
                    domestic_value = round(domestic_value)
                    domestic_value_str = str(domestic_value).strip()
                    

                
                
                trn_ref_value = row['TRN REFERENCE']
                if pd.isna(trn_ref_value):
                    trn_ref_str = "Null"  # Set a default value or handle the missing value case
                else:
                    trn_ref_str = str(trn_ref_value).strip()
                
                branch_code_value = row['BRANCH CODE']
                if pd.isna(branch_code_value):
                    branch_code_str = "Null"  # Set a default value or handle the missing value case
                else:
                    branch_code_str = str(branch_code_value).strip()
                
                acc_num_value = row['ACCOUNT NUMBER']
                if pd.isna(acc_num_value):
                    acc_num_str = "Null"  # Set a default value or handle the missing value case
                else:
                    acc_num_str = str(acc_num_value).strip()
                    acc_num_str = acc_num_str.strip('.0')
                
                id_num_value = row['ID NUMBER ']
                if pd.isna(id_num_value):
                    id_num_str = "Null"  # Set a default value or handle the missing value case
                else:
                    id_num_str = str(id_num_value).strip()
                
                tel_num_value = row['TELEPHONE']
                if pd.isna(tel_num_value):
                    tel_num_str = "Null"  # Set a default value or handle the missing value case
                else:
                    tel_num_str = str(tel_num_value).strip()
                    tel_num_str = tel_num_str.strip('.0')

                
                
                
                non_res_name = row['NON RES NAME '].strip() if not pd.isna(row['NON RES NAME ']) else "Null"
                non_res_name = non_res_name.split()[0]


                res_name = row['RES NAME '].strip() if not pd.isna(row['RES NAME ']) else "Null"
                res_name = res_name.split()[0] if not res_name == "" else "Null"

                bob_category = row['BOP CATEGORY '].strip() if not pd.isna(row['BOP CATEGORY ']) else "Null"
                bob_category = bob_category.replace("/", "")
                
                original_transaction.attrib = {
                    'LineNumber': line_number,
                    'ReportingQualifier': row['REPORTING QUALIFIER '].strip() if not pd.isna(row['REPORTING QUALIFIER ']) else "Null",
                    'Flow': row['FLOW'].strip()  if not pd.isna(row['FLOW']) else "Null",
                    'ReplacementTransaction': row['REPLACEMENT TRANSACTION'].strip() if not pd.isna(row['REPLACEMENT TRANSACTION']) else "Null",
                    'ValueDate': value_date_str,
                    'TrnReference': trn_ref_str,
                    'BranchCode': branch_code_str,
                    'BranchName': row['BRANCH NAME'].strip() if not pd.isna(row['BRANCH NAME']) else "Null",
                    'OriginatingCountry': row['ORIGINATING COUNTRY '].strip() if not pd.isna(row['ORIGINATING COUNTRY ']) else "Null",
                    'ReceivingBank': row['RECEIVING BANK'].strip() if not pd.isna(row['RECEIVING BANK']) else "Null",
                    'ReceivingCountry': row['RECEIVING COUNTRY'].strip() if not pd.isna(row['RECEIVING COUNTRY']) else "Null",
                    'TotalValue': total_value_str
                }

                non_resident = ET.SubElement(original_transaction, 'NonResident')
                individual = ET.SubElement(non_resident, 'Individual')
                individual.attrib = {
                    'Surname': row['NONRES SURNAME'].strip() if not pd.isna(row['NONRES SURNAME']) else "Null",
                    'Name': non_res_name,
                    'Gender': row['NON RES GENDER'].strip() if not pd.isna(row['NON RES GENDER']) else "M"
                }
                additional_non_resident_data = ET.SubElement(individual, 'AdditionalNonResidentData')
                additional_non_resident_data.attrib = {
                    'Country': row['NON RES COUNTRY'].strip() if not pd.isna(row['NON RES COUNTRY']) else "Null",
                    'AccountNumber': row['NON RES ACCOUNT NUMBER'].strip() if not pd.isna(row['NON RES ACCOUNT NUMBER']) else "Null"
                }

                resident_customer_account_holder = ET.SubElement(original_transaction, 'ResidentCustomerAccountHolder')
                individual_customer = ET.SubElement(resident_customer_account_holder, 'IndividualCustomer')
                individual_customer.attrib = {
                    'Surname': row['RES SURNAME '].strip() if not pd.isna(row['RES SURNAME ']) else "Null",
                    'Name': res_name,
                    'Gender': row['GENDER'].strip() if not pd.isna(row['GENDER']) else "M",
                    'DateOfBirth': dob_str,
                    'IDNumber': id_num_str
                }
                additional_customer_data = ET.SubElement(individual_customer, 'AdditionalCustomerData')
                additional_customer_data.attrib = {
                    'AccountName': res_name,
                    'AccountIdentifier': row['ACCOUNT IDENTIFIER '].strip() if not pd.isna(row['ACCOUNT IDENTIFIER ']) else "Null",
                    'AccountNumber': acc_num_str,
                    'StreetAddressLine1': row['STREET ADDRESS LINE 1'].strip() if not pd.isna(row['STREET ADDRESS LINE 1']) else "Null",
                    'StreetSuburb': row['STREET  SU BURB'].strip() if not pd.isna(row['STREET  SU BURB']) else "Null",
                    'StreetCity': row['STREET CITY '].strip() if not pd.isna(row['STREET CITY ']) else "Null",
                    'StreetRegion': row['STREET REGION '].strip() if not pd.isna(row['STREET REGION ']) else "Null",
                    'PostalAddressLine1': row['POSTAL ADDRESS LINE 1'].strip() if not pd.isna(row['POSTAL ADDRESS LINE 1']) else "Null",
                    'PostalSuburb': row['POSTAL SUBURB'].strip() if not pd.isna(row['POSTAL SUBURB']) else "Null",
                    'PostalCity': row['POSTAL CITY'].strip() if not pd.isna(row['POSTAL CITY']) else "Null",
                    'PostalRegion': row['POSTAL REGION '].strip() if not pd.isna(row['POSTAL REGION ']) else "Null",
                    'ContactSurname': row['CONTACT SURNAME'].strip() if not pd.isna(row['CONTACT SURNAME']) else "Null",
                    'ContactName': row['CONTACT NAME'].strip() if not pd.isna(row['CONTACT NAME']) else "Null",
                    'Telephone': tel_num_str
                }

                monetary_details = ET.SubElement(original_transaction, 'MonetaryDetails')
                monetary_details.attrib = {
                    'SequenceNumber': '1',
                    'MoneyTransferAgentIndicator': row['MONEY TRANSFER INDICAT0R'].strip() if not pd.isna(row['MONEY TRANSFER INDICAT0R']) else "Null",
                    'DomesticValue': domestic_value_str,
                    'DomesticCurrencyCode': row['DOMESTIC CURRENCY CODE '].strip() if not pd.isna(row['DOMESTIC CURRENCY CODE ']) else "Null",
                    'ForeignValue': foreign_value_str,
                    'ForeignCurrencyCode': row['FOREIGN CURRENCY CODE '].strip() if not pd.isna(row['FOREIGN CURRENCY CODE ']) else "Null",
                    'BoPCategory': bob_category,
                    'LocationCountry': row['LOCATION COUNTRY '].strip() if not pd.isna(row['LOCATION COUNTRY ']) else "Null"
                }
            except Exception as e:
                print(f"Error in row {count + 1}: {e}")
                print(row)  # Print the entire row to identify the problematic column
            count +=1

        # Create the XML file
        xml_tree = ET.ElementTree(root)
        xml_file = os.path.splitext(os.path.basename(excel_path))[0] + '.xml'
        xml_path = os.path.join(output_folder, xml_file)

        # Serialize the XML tree to a string
        xml_string = ET.tostring(root)

        # Use minidom to parse and format the XML string
        parsed_xml = minidom.parseString(xml_string)
        formatted_xml = parsed_xml.toprettyxml(indent='  ')

        # Write the formatted XML to the file
        try:
            with open(xml_path, 'w') as xml_file:
                xml_file.write(formatted_xml)
                if not existing:
                    icon.notify(f"File {os.path.splitext(os.path.basename(excel_path))[0] + '.xml'} created successfully!")
                # Move the Excel file to the completed folder
                completed_path = os.path.join(completed_folder, os.path.basename(excel_path))
                shutil.move(excel_path, completed_path)

        except Exception as e:
            if not existing:
                icon.notify(f"Error writing XML file: {str(e)}")
            # Move the invalid excel to the failed folder
            failed_path = os.path.join(failed_folder, os.path.basename(excel_path))
            shutil.move(excel_path, failed_path)

    except Exception as e:
        # Log the error
        logger.error(f'Error processing Excel file: {excel_path}\n{str(e)}')


def process_existing_files():
    # Get the list of existing Excel files in the input folder
    existing_files = [file for file in os.listdir(input_folder) if file.endswith('.xlsx')]

    # Process each existing file
    for file in existing_files:
        file_path = os.path.join(input_folder, file)
        process_excel(file_path, True)


def on_quit_callback(icon, item):
    observer.stop()
    icon.stop()

def open_folder_settings():
    subprocess.run(["python", "settings.py"])
    
def create_system_tray_icon():
    icon_path = 'icon.png'
    image = Image.open(icon_path)

     # Create the QApplication instance
    
    # Create the menu and add options
    menu = (
        item('Folder Settings', open_folder_settings),
        item('Terminate', on_quit_callback)
    )
    icon = pystray.Icon("Mike", image, "STD Reporter", menu)

     # Create the file watcher for stdsettings.ini
    file_watcher = QFileSystemWatcher([os.path.join(os.getcwd(), 'stdsettings.ini')], parent=app)
    file_watcher.fileChanged.connect(show_notification)
    
    return icon

def show_notification():
    
    icon.notify("The settings file has been modified.")


if __name__ == '__main__':
    
    # Process existing files in the input folder
    process_existing_files()
     # Create the QApplication instance
    app = QApplication(["STD Reporter"])

    # Create an event handler and observer
    event_handler = FileHandler()
    observer = Observer()



    # Set up the observer to watch the input_folder
    observer.schedule(event_handler, input_folder, recursive=False)
    observer.start()

    # Create system tray icon
    icon = create_system_tray_icon()
    icon.run()


