import os
import sys
import datetime
from dateutil.relativedelta import relativedelta
import xlwings as xw
import re
import pprint
from main import AnalyzeExcel

crm_data = {}

class AnalyzeCRMExcel(AnalyzeExcel):
  def parseCRM(self):
    # Open Excel
    wb = xw.Book(self.import_file)
    sheet = wb.sheets(1)

    start_row = 2
    last_row = sheet.range('A1').current_region.last_cell.row
    # last_row = 500

    for i in range(start_row, last_row + 1):
      try:
        service_group = sheet.range(f'Y{i}').value
        team = self.find_team(sheet.range(f'Q{i}').value, service_group)
        if (team != 'CRM'):
          continue

        # Create day is key of crm_data
        create_day = sheet.range(f'D{i}').value
        key = create_day.strftime("%Y/%m/%d")

        if not key in crm_data:
          crm_data[key] = {}

        # Request or Incident
        ticket_type = sheet.range(f'N{i}').value
        category = 'Request' if ticket_type == 'Request' else "Incident"

        if not category in crm_data[key]:
          crm_data[key][category] = 1
        else:
          crm_data[key][category] += 1
        
      except Exception as e:
        print(e)
        import traceback
        traceback.print_exc()
    
    # wb.close()
    wb.app.quit()

class CRMReport:
  def __init__(self):
    self.output_dir = os.path.join(os.getcwd(), 'crm')
    self.output_file = os.path.join(self.output_dir, 'crm_output.xlsx')

  def dumpCRM(self):
    # Open Excel
    wb = xw.Book()
    sheet = wb.sheets(1)

    sheet.range('B2').value = 'Create Date'
    sheet.range('C2').value = 'Incident'
    sheet.range('D2').value = 'Request'

    cur_position = sheet.range('B3')

    i = 0
    for k, v in crm_data.items():
      cur_position.offset(i, 0).value = k
      cur_position.offset(i, 1).value = v['Incident']
      cur_position.offset(i, 2).value = v['Request']
      i += 1

    wb.save(self.output_file)
    wb.app.quit()

analyzeCRMExcel = AnalyzeCRMExcel("AK_Ticket_2019-2.xlsx")
analyzeCRMExcel.parseCRM()
pprint.pprint(crm_data)

crmReport = CRMReport()
crmReport.dumpCRM()
