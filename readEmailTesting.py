import os
import imaplib
import email
import traceback
from datetime import datetime
from email.header import decode_header
import pandas as pd
from bs4 import BeautifulSoup
import time
import glob
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

def clean(text):  # TO create a folder
    return "".join(c if c.isalnum() else "_" for c in text)

def sort_list(List):
    List.sort(key=lambda l: l[2])
    return List

# Export to different sheets and loop the header
def excel_export_daily(filename, excelOutput):
    df = pd.DataFrame(filename)
    writer = pd.ExcelWriter(excelOutput)
    df.to_excel(writer, sheet_name=excelOutput.split(".")[0], header=False, index=False)
    writer.close()
    print(excelOutput + " Export Done")

def excel_export_monthly(filename, excelOutput):
    filename_copy = filename
    row_header = filename_copy.pop(0)
    df = pd.DataFrame(sort_list(filename_copy))
    writer = pd.ExcelWriter(excelOutput)
    df.to_excel(writer, sheet_name=excelOutput.split(".")[0], startrow=1, header=False, index=False)
    worksheet = writer.sheets['Daily Sorted']
    for x in range(len(row_header)):
        worksheet.write(0, x, row_header[x])
    writer.close()
    print(excelOutput + " Export Done")

def excel_monthly_sum(filename, header_lists, excelOutput):
    filename_copy = filename
    df = pd.DataFrame(filename_copy, columns=header_lists)
    # Since Transaction Date is considered as an Object, we convert it back to Datetime
    df['Transaction Date'] = pd.to_datetime(df['Transaction Date'], format='%Y-%m-%d')
    df['year'] = df['Transaction Date'].dt.year
    df['month'] = df['Transaction Date'].dt.month
    df_grouped = df.groupby([df['year'], df['month']]).sum()
    writer = pd.ExcelWriter(excelOutput)
    df_grouped.to_excel(writer, sheet_name=excelOutput.split(".")[0])
    writer.close()

def merge_excel(excelFiles, excelOutput):
    with pd.ExcelWriter(excelOutput) as writer:
        for excel in excelFiles:  # For each excel
            sheet_name = pd.ExcelFile(excel).sheet_names[0]  # Find the sheet name
            df = pd.read_excel(excel)  # Create a dataframe
            df.to_excel(writer, sheet_name=sheet_name, index=False)  # Write it to a sheet in the output excel
    print("Excel Files Merged")

def read_email_from_outlook(FROM_EMAIL, FROM_PWD, SMTP_SERVER, SMTP_PORT):
    try:
        mail = imaplib.IMAP4_SSL(SMTP_SERVER)
        mail.login(FROM_EMAIL, FROM_PWD)
        status, messages = mail.select("DAILY")
        # total number of emails
        messages = int(messages[0])
        # print(messages)

        data = mail.search(None, 'ALL')
        mail_ids = data[0]
        id_list = mail_ids.split()

        # Export to Excel
        rows, cols = (messages + 1, 9)
        row_list = [[0 for i in range(cols)] for j in range(rows)]
        col_remove = 0  # To remove email output with no table

        # Fill in the header
        header_lists = ["A/C No", "A/C Name", "Transaction Date", "Closed P/L", "Floating P/L", "Previous Ledger Balance", "Balance", "Equity", "Available Margin"]
        for x in range(len(header_lists)):
            row_list[0][x] = header_lists[x]

        # Counter
        excelCounter = 1

        # messages+1 to include the Header row
        # for i in range(messages, (messages+1) - (messages+1), -1):
        for i in range(1, messages + 1):
            type, data = mail.fetch(str(i), '(RFC822)')
            for response in data:
                if isinstance(response, tuple):
                    # parse a bytes email into a message object
                    msg = email.message_from_bytes(response[1])
                    # decode the email subject
                    subject, encoding = decode_header(msg["Subject"])[0]
                    if isinstance(subject, bytes):
                        # if it's a bytes, decode to str
                        subject = subject.decode(encoding)
                    # decode email sender
                    From, encoding = decode_header(msg.get("From"))[0]
                    # if isinstance(From, bytes):
                    #    From = From.decode(encoding)
                    # print("Subject:", subject)
                    # print("From:", From)
                    if msg.is_multipart():
                        for part in msg.walk():
                            # extract content type of email
                            content_type = part.get_content_type()
                            content_disposition = str(part.get("Content-Disposition"))
                            try:
                                # get the email body
                                body = part.get_payload(decode=True).decode()
                            except:
                                pass
                            if content_type == "text/plain" and "attachment" not in content_disposition:
                                # print text/plain emails and skip attachments
                                print(body)
                            elif "attachment" in content_disposition:
                                # download attachment
                                filename = part.get_filename()
                                if filename:
                                    folder_name = clean(subject)
                                    if not os.path.isdir(folder_name):
                                        # make a folder for this email (named after the subject)
                                        os.mkdir(folder_name)
                                    filepath = os.path.join(folder_name, filename)
                                    # download attachment and save it
                                    open(filepath, "wb").write(part.get_payload(decode=True))
                    else:
                        # extract content type of email
                        content_type = msg.get_content_type()
                        # get the email body
                        body = msg.get_payload(decode=True).decode()
                        if content_type == "text/html":
                            # print only text email parts
                            # print(body)

                            soup = BeautifulSoup(body, 'lxml')  # Parse the HTML as a string
                            table = soup.table

                            if table is None:
                                col_remove += 1
                                continue
                            else:
                                rows = table.find_all('tr')
                                data = []
                                for tr in rows:
                                    td = tr.find_all('td')
                                    data.append([i.text for i in td])

                                acNo = data[0][0].split(":")[1].strip()
                                acName = data[0][1].split(":")[1].strip()
                                txnDate_temp = data[0][3].split(",")[0].strip()
                                txnDate = datetime.strptime(txnDate_temp, '%Y %B %d').date()
                                closedPl = float(data[len(data) - 5][1].strip().replace(" ", ""))
                                floatingPl = float(data[len(data) - 6][4].strip().replace(" ", ""))
                                prevLedgerBal = float(data[len(data) - 6][1].strip().replace(" ", ""))
                                balance = float(data[len(data) - 3][1].strip().replace(" ", ""))
                                equity = float(data[len(data) - 4][4].strip().replace(" ", ""))
                                availMargin = float(data[len(data) - 2][1].strip().replace(" ", ""))

                                print(excelCounter)
                                # Format adjustable
                                print("A/C No: " + acNo)
                                print("A/C Name: " + acName)
                                print("Transaction Date: " + txnDate.strftime('%Y %B %d'))
                                print("Closed P/L: " + str(closedPl))
                                print("Floating P/L: " + str(floatingPl))
                                print("Previous Ledger Balance: " + str(prevLedgerBal))
                                print("Balance: " + str(balance))
                                print("Equity: " + str(equity))
                                print("Available Margin: " + str(availMargin) + '\n')

                                # Fill Excel row
                                row_list[excelCounter][0] = acNo
                                row_list[excelCounter][1] = acName
                                row_list[excelCounter][2] = txnDate
                                row_list[excelCounter][3] = closedPl
                                row_list[excelCounter][4] = floatingPl
                                row_list[excelCounter][5] = prevLedgerBal
                                row_list[excelCounter][6] = balance
                                row_list[excelCounter][7] = equity
                                row_list[excelCounter][8] = availMargin

                                excelCounter += 1

        row_list = row_list[:-col_remove]
        print(len(row_list))
        excel_export_daily(row_list, 'sampleOutputDaily.xlsx')
        excel_export_monthly(row_list, 'sampleOutputSorted.xlsx')
        excel_monthly_sum(row_list, header_lists, 'sampleOutputMonthlySummary.xlsx')
        excel_files = ['sampleOutputDaily.xlsx', 'sampleOutputSorted.xlsx', 'sampleOutputMonthlySummary.xlsx']
        merge_excel(excel_files, 'sampleOutputFinal.xlsx')

        mail.close()
        mail.logout()

    except Exception as e:
        traceback.print_exc()
        print(str(e))


def main():
    # -------------------------------------------------
    # Confidential Information
    # ------------------------------------------------
    FROM_EMAIL = "xxx@outlook.com"
    FROM_PWD = "abcde123"
    SMTP_SERVER = "outlook.office365.com"
    SMTP_PORT = 993

    read_email_from_outlook(FROM_EMAIL, FROM_PWD, SMTP_SERVER, SMTP_PORT)

if __name__ == "__main__":
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))
