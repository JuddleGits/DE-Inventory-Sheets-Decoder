import os
import tkinter as tk
import tkinter.messagebox as msg
from Google import Create_Service

class SheetsApp(tk.Frame):

    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        self.text = tk.Label(self, text="DE Inventory Decoder", font=("Clarendon Blk BT Black", 16))
        self.text.pack()
        self.text2 = tk.Label(self, text="Insert Spreadsheet ID Below", font=("Consolas", 12))
        self.text2.pack(pady=10)
        self.entry = tk.StringVar()
        self.text_field = tk.Entry(self, width=30, textvariable=self.entry)
        self.text_field.pack(pady=10)
        self.button = tk.Button(self, text="Run Program", height=3, command=self.create_service)
        self.button.pack(pady=10)
        self.quit_button = tk.Button(self, text="Quit Program", height=3, command=self.master.destroy)
        self.quit_button.pack(pady=10)
        self.in_progress = tk.Label(self, text="Decoding In Progress...", font=("Clarendon Blk BT Black", 12))
        self.finished = tk.Label(self, text="Decoding Has Finished!", font=("Clarendon Blk BT Black", 12))

    def create_service(self):
        spreadsheetID = self.entry.get()
        if len(spreadsheetID) == 0:
            msg.showerror(title="ERROR!", message="The SpreadsheetID field is empty!")
            return

        self.finished.pack_forget()
        self.in_progress.pack()
        self.master.update()

        """
		Can get credentials.json file from https://developers.google.com/sheets/api/quickstart/python
		"""
        CLIENT_SECRET_FILE = 'credentials.json'
        API_NAME = 'sheets'
        API_VERSION = 'v4'
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

        service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)
        self.cycle_sheets(service, spreadsheetID)

    def cycle_sheets(self, service, spreadsheetID):
        try:
        	res = service.spreadsheets().get(
        	spreadsheetId=spreadsheetID,
        	fields='sheets(data/rowData/values/userEnteredValue,properties(index,sheetId,title))'
        	).execute()
        except Exception:
            msg.showerror(title="ERROR!", message="The SpreadsheetID given does not match any spreadsheet!")
            self.in_progress.pack_forget()
            return
        
        sheet_index = 0

        while sheet_index < len(res['sheets']):
            self.sheet_decode(sheet_index, res, service, spreadsheetID)
            sheet_index += 1

    def sheet_decode(self, sheet_index, res, service, spreadsheetID):
        sheet_name = res['sheets'][sheet_index]['properties']['title']
        last_row = len(res['sheets'][sheet_index]['data'][0]['rowData'])
        cell_range_read = 'C2:C' + str(last_row)
        cell_range_read2 = 'A2:A' + str(last_row)
        cell_range_insert = 'D2'

        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheetID,
            range=sheet_name + '!' + cell_range_read
            ).execute()

        result2 = service.spreadsheets().values().get(
            spreadsheetId=spreadsheetID,
            range=sheet_name + '!' + cell_range_read2
            ).execute()

        temp = 0
        strNum = ""
        decodeVals = ()

        for value in result['values']:
            try:
                numStr = value[0]
            except IndexError:
                decodeVals += tuple(['0'])
                continue
            if numStr.isdigit():
                for digit in numStr:
                    num = int(digit)
                    if num != 0:
                        temp = 10 - num
                        strNum += str(temp)
                    else:
                        strNum += str(num)
                if int(strNum) >= 100:
                    strNum = strNum[:len(strNum) - 2] + '.' + strNum[len(strNum) - 2:]
                elif int(strNum) >= 10:
                    strNum = '0.' + strNum
                else:
                    strNum = '0.0' + strNum
                # print(strNum)
                decodeVals += tuple([str(strNum)])
                strNum = ""
            else:
                decodeVals += tuple(['0'])

        quantities = ()
        totals = ()
        for value in result2['values']:
            numStr = value[0]
            if numStr.isdigit():
                quantities += tuple([str(numStr)])
            else:
                quantities += tuple(['0'])

        n = 0
        while n < len(quantities):
            totalPrice = int(quantities[n]) * float(decodeVals[n])
            totals += tuple([str(totalPrice)])
            n += 1

        finalValues = (decodeVals, totals)
        value_range_body = {
            'majorDimension': 'COLUMNS',
            'values': finalValues
        }

        service.spreadsheets().values().update(
            spreadsheetId=spreadsheetID,
            valueInputOption='USER_ENTERED',
            range=sheet_name + '!' + cell_range_insert,
            body=value_range_body
        ).execute()

        self.in_progress.pack_forget()
        self.finished.pack()
