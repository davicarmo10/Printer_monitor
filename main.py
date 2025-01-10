import win32print
import time
from openpyxl import Workbook, load_workbook
from datetime import datetime

register_path = r"\\Path\to\save\logs.xslx"
printer_name = "Printer Name"
rather_document = "Fast Report Document"

# Function to initialize the Excel file
def init_excel():
    try:
        # try to load the file
        workbook = load_workbook(register_path)
        sheet = workbook.active
    except FileNotFoundError:
        # if doesnt exist, create a new file
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Logs of Print"
        sheet.append(["Date/time", "ID", "Document", "User", "Status"])
        workbook.save(register_path)
    return workbook, sheet

# Function to read the IDs of the jobs already saved in the Excel file
def load_id(sheet):
    add_ids = set()
    for row in sheet.iter_rows(min_row=2, max_col=2, values_only=True):
        add_ids.add(row[0])  # Adicionar o ID do trabalho na lista
    return add_ids

# Open conection with printer and monitor

def printer_monitor (printer_name_in):
    
    try:
        nPrinter = win32print.OpenPrinter(printer_name_in)
        workbook, sheet = init_excel()
        ids = load_id(sheet)

        # Get info about the printer

        while True:
            print("-----------------------Waiting job--------------------------")
            jobs = win32print.EnumJobs(nPrinter, 0, -1, 2)  
            if jobs:
                for job in jobs: 
                    job_id = job['JobId']
                    document = job['pDocument']
                    
                    if document == rather_document:
                        if job_id not in ids: #verifica ids
                            # Get infos of the print
                            date_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            document = job['pDocument']
                            user = job['pUserName']
                            status = job.get('Status', 'Unknown')

                            #Show in console
                            print(f"\nPrinted pages: {job['PagesPrinted']}")
                            print(f"Job ID: {job['JobId']}")
                            print(f"Status: {job['Status']}")
                            print(f"Document: {job['pDocument']}\n")

                            # Save document in Excel file
                            sheet.append([date_time, job_id, document, user, status])
                            ids.add(job_id)
                            workbook.save(register_path)    
            else:
                print(f"No jobs in a row '{printer_name_in}'.")
            time.sleep(5)
           
    except Exception as e:
        print(f"error")

# close conection with printer
    finally:
        win32print.ClosePrinter(nPrinter)

if __name__ == '__main__':
    printer_monitor(printer_name)
