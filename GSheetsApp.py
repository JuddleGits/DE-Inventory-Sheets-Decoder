import os
import tkinter as tk
from Google import Create_Service

class SheetsApp(tk.Frame):
	def __init__(self, master=None):
		super().__init__(master)
		self.master = master
		self.pack()
		self.create_button()

	def create_button(self):
		self.button = tk.Button(self)
		self.button.place(relx=0.5, rely=0.5, anchor="center")
		self.button["text"] = "Run Program"
		self.button["command"] = self.create_service
		self.button.pack()

	def create_service(self):
		"""
		Can retrieve your credentials.json file from https://developers.google.com/sheets/api/quickstart/python
		"""
		CLIENT_SECRET_FILE = 'credentials.json' 
		API_NAME = 'sheets'
		API_VERSION = 'v4'
		SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

		service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)
		spreadsheetID = '1mIqhz62uyh1-W3n1JaCfzp6F9WVwhD92JAA1DZJOEmc'
		self.cycle_sheets(service, spreadsheetID)

	def cycle_sheets(self, service, spreadsheetID):
		res = service.spreadsheets().get(
			spreadsheetId=spreadsheetID,
			fields='sheets(data/rowData/values/userEnteredValue,properties(index,sheetId,title))'
			).execute()

		sheet_index = 0

		while sheet_index < len(res['sheets']):
			self.sheet_decode(sheet_index, res, service, spreadsheetID)
			sheet_index += 1

	def sheet_decode(self, sheet_index, res, service, spreadsheetID):
		sheet_name = res['sheets'][sheet_index]['properties']['title']
		last_row = len(res['sheets'][sheet_index]['data'][0]['rowData'])
		cell_range_read = 'C1:C' + str(last_row)
		cell_range_read2 = 'A1:A' + str(last_row)
		cell_range_insert = 'D1'

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
			numStr = value[0]
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
				print(strNum)
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
