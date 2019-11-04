import os
import sys
import datetime
from dateutil.relativedelta import relativedelta
import xlwings as xw
import re
import pprint
import numpy as np
import matplotlib.pyplot as plt
import time

cwd = os.getcwd()
# 作業は月初に行うが、データのimport先とexport先は先月になる
last_month_date = datetime.date.today() - relativedelta(months=1)
last_month = last_month_date.strftime("%Y%m")

draw_data = {
}
track_month = 18
today = datetime.date.today()
for i in reversed(range(1, track_month+1)):
  month = today - relativedelta(months=i)
  ticket_date_key = month.strftime("%Y%m")
  draw_data[ticket_date_key] = {
    "overview": {
      "incident": 0,
      "request": 0,
      "total" : 0,
    },
    "incident": {
      "inquiry": 0,
      "problem": 0,
      "total": 0,
    },
    "request": {
      "id": 0,
      "other": 0,
      "total": 0
    },
    "open": {
      "shortopen": 0,
      "longopen": 0,
      "total": 0
    }
  }

class AnalyzeDrawData:
  def __init__(self):
    self.draw_data = draw_data
    self.output_dir = os.path.join(cwd, 'output', last_month)
    self.output_file = os.path.join(self.output_dir, 'output.xlsx')
    self.output_longopen_dir = os.path.join(cwd, 'output', 'longopen')
    self.output_longopen_file = os.path.join(
      self.output_longopen_dir, 'longopen_history.xlsx')

  def readExcel(self):
    wb_output = xw.Book(self.output_file)
    sheet_output = wb_output.sheets(1)

    cur_postion = sheet_output.range('c2')
    
    while (1):
      if str(int(cur_postion.value)) in self.draw_data:
        break

      cur_postion = cur_postion.offset(0, 1)

    while (1):
      cur_month = str(int(cur_postion.value))
      
      self.draw_data[cur_month]['overview']['incident'] = int(cur_postion.offset(1,0).value)
      self.draw_data[cur_month]['overview']['request']  = int(cur_postion.offset(2,0).value)
      self.draw_data[cur_month]['overview']['total']    = int(cur_postion.offset(3,0).value)
      self.draw_data[cur_month]['incident']['inquiry']  = int(cur_postion.offset(6,0).value)
      self.draw_data[cur_month]['incident']['problem']  = int(cur_postion.offset(7,0).value)
      self.draw_data[cur_month]['incident']['total']    = int(cur_postion.offset(8,0).value)
      self.draw_data[cur_month]['request']['id']        = int(cur_postion.offset(11,0).value)
      self.draw_data[cur_month]['request']['other']     = int(cur_postion.offset(12,0).value)
      self.draw_data[cur_month]['request']['total']     = int(cur_postion.offset(13,0).value)
   
      cur_postion = cur_postion.offset(0, 1)
      if cur_postion.value == None:
        break

  def readOpen(self):
    wb_longopen = xw.Book(self.output_longopen_file)  
    sheet_longopen = wb_longopen.sheets('overview')

    cur_postion = sheet_longopen.range('c2')
    
    
    while (1):
      if str(int(cur_postion.value)) in self.draw_data:
        break

      cur_postion = cur_postion.offset(0, 1)

    while (1):
      cur_month = str(int(cur_postion.value))
      
      self.draw_data[cur_month]['open']['shortopen'] = int(cur_postion.offset(1,0).value)
      self.draw_data[cur_month]['open']['longopen']  = int(cur_postion.offset(2,0).value)
      self.draw_data[cur_month]['open']['total']     = int(cur_postion.offset(3,0).value)
   
      cur_postion = cur_postion.offset(0, 1)
      if cur_postion.value == None:
        break

class Draw:
  def __init__(self):
    self.draw_data = draw_data
    self.output_dir = os.path.join(cwd, 'output', last_month)

  def drawAll(self):
    ary_month = []
    ary_overview_incident = []
    ary_overview_request = []
    ary_overview_total = []
    ary_incident_inquiry = []
    ary_incident_problem = []
    ary_incident_total = []
    ary_request_id = []
    ary_request_other = []
    ary_request_total = []
    ary_open_shortopen = []
    ary_open_longopen = []
    ary_open_total = []

    for k, v in draw_data.items():
      ary_month.append(f'{k[4:6]}\n{k[0:4]}')
      ary_overview_incident.append(v['overview']['incident'])
      ary_overview_request.append(v['overview']['request'])
      ary_overview_total.append(v['overview']['total'])
      ary_incident_inquiry.append(v['incident']['inquiry'])
      ary_incident_problem.append(v['incident']['problem'])
      ary_incident_total.append(v['incident']['total'])
      ary_request_id.append(v['request']['id'])
      ary_request_other.append(v['request']['other'])
      ary_request_total.append(v['request']['total'])
      ary_open_shortopen.append(v['open']['shortopen'])
      ary_open_longopen.append(v['open']['longopen'])
      ary_open_total.append(v['open']['total'])

    self.plot(ary_month, ary_overview_request,
              ary_overview_incident, 'Overview', 'Request', 'Incident')
    
    self.plot(ary_month, ary_incident_problem, ary_incident_inquiry,
              'Incident', 'Problem', 'Inqruiry')
    
    self.plot(ary_month, ary_request_other, ary_request_id,
              'Request', 'Other', 'ID')
    
    self.plot(ary_month, ary_open_shortopen,
              ary_open_longopen, 'Open', 'shortopen', 'Longopen')

  def plot(self, ary1, ary2, ary3, title, s1, s2):
    plt.figure(figsize=(12,4))
    left = np.array(ary1)
    height1 = np.array(ary2)
    height2 = np.array(ary3)

    p1 = plt.bar(left, height1, color="green")
    p2 = plt.bar(left, height2, bottom=height1, color="orange")

    # Input number in bar graph
    for x, y1, y2 in zip(left, height1, height2):
      plt.text(x, y1/2, str(y1), horizontalalignment="center")
      plt.text(x, y1+y2/2, str(y2), horizontalalignment="center")
      plt.text(x, y1+y2, str(y1+y2), horizontalalignment="center")

    plt.legend((p2[0], p1[0]), (s2, s1))

    output_file = os.path.join(self.output_dir, f'{title}.png')
    plt.savefig(output_file)

# Input data
analyzeDrawData = AnalyzeDrawData()
analyzeDrawData.readExcel()
analyzeDrawData.readOpen()

pprint.pprint(draw_data)

draw = Draw()
draw.drawAll()
