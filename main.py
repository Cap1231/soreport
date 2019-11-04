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
all_cism = {}
apps = ['CRM', 'Autoline', 'SWT', 'MmcR', 'Dealer Portal',
        'SPO', 'SAP', 'UCP', 'DH', 'OA', 'SABA', 'VPC', 'OTHER', 'DAW', 'Paragon',
        'Salestouch', 'Server', 'Network', 'Security', 'EVA']

# 作業は月初に行うが、データのimport先とexport先は先月になる
last_month_date = datetime.date.today() - relativedelta(months=1)
last_month = last_month_date.strftime("%Y%m")

# output するフォルダを作成
output_dir_path = os.path.join(cwd, 'output', '201910')
os.makedirs(output_dir_path, exist_ok=True)

# チケットのトラッキングは18ケ月(1年半年)行う
track_month = 18


class AnalyzeExcel:
  def __init__(self, file_name):
    self.import_dir = os.path.join(cwd, 'import', last_month)
    self.import_file = os.path.join(self.import_dir, file_name)
    # self.wb = xw.Book(self.import_file)

  def find_team(self, description, service_group):
    try:
      app_desc = re.search(r'\[.*\]', description).group()
    except:
      app_desc = "none"

    app_desc = app_desc.upper()

    if 'CRM' in app_desc:
      return 'CRM'
    elif 'AUTOLINE' in app_desc:
      return 'Autoline'
    elif 'MMCR' in app_desc or 'MERCEDES ME CONNECT RETAIL' in app_desc:
      return 'MmcR'
    elif 'DP' in app_desc:
      return 'Dealer Portal'
    elif 'SPO' in app_desc:
      return 'SPO'
    elif 'SAP' in app_desc:
      return 'SAP'
    elif 'UCP' in app_desc:
      return 'UCP'
    elif 'DH' in app_desc:
      return 'DH'
    elif 'EVA' in app_desc:
      return 'EVA'
    elif 'SABA' in app_desc:
      return 'SABA'
    elif 'SWT' in app_desc:
      return 'SWT'
    elif 'VPC' in app_desc:
      return 'VPC'
    elif 'DAW' in app_desc:
      return 'DAW'
    elif 'PARAGON' in app_desc:
      return 'Paragon'
    elif 'STOUCH' in app_desc:
      return 'Salestouch'
    else:
      oa_sg = [
        'ITM_O_TL_COC_FS_APPVANTAGE',
        'DCL_MBAFC_UAM_UHD',
        'FITO_EXT_FUHD_VIPSUPPORT',
        'GOA_AD_AccountMgmt',
         'GOA_AVM_2ndLevel',
         'GOA_dw-Browser_Operating',
         'GOA_dw-client_2ndLevel_ext',
         'GOA_dw-client_3rdLevel_ext',
         'GOA_dw-collab_2ndLevel_IN',
         'GOA_dw-mobile_2ndLevel_ext',
         'GOA_dw-mobile_2ndLevel_in',
         'GOA_GCS_AccessManagement',
         'GOA_GCS_ClientM_3rd',
         'GOA_GCS_ClientMaintenance',
         'GOA_GCS_ClientSettings',
         'GOA_GCS_ProblemManagement',
         'GOA_GCS_SoftwareManagement',
         'GOA_GCS_WOC',
         'GOA_GESM_2ndLevel',
         'GOA_GESM_3rdLevel',
         'GOA_PCP4_SUPPORT',
         'GOA_SDA_2ndLevel',
         'GOA_SDA_3rdLevel',
         'GOA_SMI_2ndLevel',
         'GOSC_IUHD_FL_Servicedesk',
         'InfoSec_Incident_Response',
         'ITM_O_RUN_CHANGE_IMPL_VOICE',
         'ITM_O_RUN_CHANGE_IMPL_WIN',
         'ITM_O_RUN_CHANGE_IMPL_WINPATCH',
         'ITM_O_SL_ITI_OA_CITRIX',
         'ITM_O_SL_ITI_OA_SMS',
         'ITM_O_SL_RUN_OA_CITRIX',
         'ITM_O_SL_RUN_OA_ITSPS',
         'ITM_O_SL_RUN_OA_OFFICE_CLIENT',
         'MBJ_OPC_SL',
         'MBJ_SL_OA',
         'MBJ_SL_PCS',
         'PVMGMT_COLL_OM_DAI',
         'PVMGMT_COLL_OPERATION',
         'PVMGMT_COLL_SLS_TSI',
         'PVMGMT_eCSLS',
         'S_CUHD_Collaboration',
         'S_CUHD_IPAD_IPHONE',
         'S_CUHD_IT-SPS_2nd',
         'S_CUHD_mobile-computing',
         'S_CUHD_Notes',
         'S_CUHD_SSL_VPN',
         'S_CUHD_SW_Cube',
         'S_CUHD_WhoIsWho',
         'CDCSP_KCS_SERVICE',
         'GOA_AD_Request-Mgmt_OFF',
         'ITM_O_ITI_CHANGE_IMPL_DB'
      ]
      swt_sg = [
        'ITM_O_SL_CoC_SWT_FICO',
        'MBCAN_NA_MBDev',
        'CG_GO_DISY',
        'CoFiCo_FL_PO_Finance',
         'ITM_O_SL_CoC_SWT_AUTH',
         'ITM_O_SL_COC_SWT_Demand_FICO',
         'ITM_O_SL_COC_SWT_Demand_LO',
         'ITM_O_SL_CoC_SWT_GRC',
         'ITM_O_SL_CoC_SWT_LO',
         'ITM_O_SL_CoC_SWT_SD',
         'ITM_O_SL_CoC_SWT_XI',
         'ITM_O_TL_CoC_SWT_DEMAND',
         'ITM_O_TL_CoC_SWT_DEMAND_FICO',
         'ITM_O_TL_CoC_SWT_DEMAND_LO',
         'ITM_O_TL_CoC_SWT_DEMAND_SD',
         'ITM_O_TL_CoC_SWT_LO',
         'ITM_O_TL_SDC_SWT_FICO',
         'ITM_O_TL_SDC_SWT_SD',
         'ITM_O_TL_SDC_SWT_XI',
         'MBJ_SL_AMS_SWT_FICO',
         'MBJ_SL_AMS_SWT_Vehicle',
         'MBJ_SL_Finance',
         'MBJ_SL_MBJ_SWT_FICO',
         'MBJ_SL_MBJ_SWT_Vehicle',
         'MBJ_SL_Vehicle',
         'ITM_O_SL_SDC_SWT_LO'
      ]
      crm_sg = [
        'ITM_O_TL_COC_CRM_NM',
        'ITM_O_SL_COC_CDM',
        'ITM_O_SL_COC_CDM_NM',
        'ITM_O_SL_COC_CRM',
         'ITM_O_SL_COC_CRM_NM',
         'ITM_O_SL_COC_MC',
         'ITM_O_TL_COC_CRM',
         'MBJ_FL_CRM',
         'MBJ_SL_CRM'
      ]
      dp_sg = [
        'MBJ_SL_AMS_DealerPortal',
        'MBJ_SL_DealerPortal',
        'MBJ_SL_MBJ_DealerPortal'
      ]
      daw_sg = [
        'MBJ_FL_IDM',
        'MBJ_SL_AMS_DAW',
        'MBJ_SL_DAW',
        'MBJ_SL_IDM',
         'MBJ_SL_MBJ_DAW'
      ]
      dsd_sg = [
        'MBJ_FL_Dealer_SD'
      ]
      pg_sg = [
        'MBJ_SL_Aftersales',
        'MBJ_SL_AMS_PARAGON',
        'MBJ_SL_MBJ_PARAGON'
      ]
      nw_sg = [
        'MBJ_SL_NO',
        'MBJ_FL_Se',
        'ITM_O_SL_',
        'MBJ_SL_NO',
        'MBJ_SL_FS',
        'MBJ_SL_CR',
        'MBJ_SL_TM',
        'ITM_O_SL_',
        'ITI_SServ',
        'ITM_O_FL_ServiceDesk',
        'MBJ_OPC_FL',
        'MTC_ITS_Infrastructure',
        'ITM_O_SL_ITI_NETWORK_ADC',
        'ITM_O_RUN_CHANGE_IMPL_NS',
        'ITM_O_RUN_CHANGE_IMPL_NETSEC',
        'MBJ_FL_Dealer_SD',
        'ITM_O_ITI_RUN_CHANGE_IMPL_CO',
        'ITM_O_SL_ITI_NETWORK_RTC',
        'ITM_O_SL_ITI_Japan_ServerOps',
        'ITM_O_SL_RUN_NETSEC',
        'DCL_Infrastructure_Nework',
        'ITM_O_SL_RUN_VIDEO',
        'ITM_O_SL_RUN_AUTOLINE',
        'ITM_O_SL_ITI_FIREWALL',
        'ITM_O_FL_RUN_NS_OPS',
      ]
      server_sg = [
        'MBJ_SL_AML',
        'MBJ_SL_DO',
        'SECURITY_REVIEW_EPA',
        'MBJ_SL_TML',
        'MBJ_OPC_FL',
        'CISM_Deployment_int',
        'ITM_O_SL_RUN_SERVER_UNIX',
        'MBJ_SL_FS_Group1',
        'ITM_O_FL_ServiceDesk',
        'ITM_O_ITI_CHANGE_CO',
        'DCL_Infrastructure_SL_Server',
        'MBJ_SL_FSPAC',
        'ITM_O_RUN_CHANGE_IMPL_UNIX',
        'ITM_O_RUN_CHANGE_IMPL_MW',
        'SWT_TSS_TL_TSS',
        'ITM_O_ITI_RUN_INCIDENT_CO',
        'ITM_O_RUN_CHANGE_IMPL_STORAGE',
        'ITM_O_SL_RUN_DATABASE',
        'ITM_O_SL_RUN_BACKUP',
        'ITM_O_SL_ITI_MONITORING',
        'extMF_Supp_Host_Change',
        'ITM_O_SL_RUN_SWT_AUTH',
        'ITM_O_SL_ITI_SAP_BASIS',
        'ITM_O_SL_ITI_SERVER_UNIX',
        'ITM_O_SL_COC_SWT_Demand_SD',
        'ITM_O_SL_RUN_VM',
        'ITM_O_RUN_CHANGE_IMPL_WEB',
        'ITM_O_ITI_RUN_CHANGE_IMPL_CO',
        'ITM_O_RUN_CHANGE_IMPL_AUTO',
        'ITM_O_TL_ITI_DATABASE',
        'ITM_O_SL_RUN_VOICE',
        'CISM_3rdLevel',
        'GOA_ESMO_L3',
        'ITM_O_RUN_CHANGE_IMPL_DB',
        'PVMGMT_COLL_PPSP_TSI',
        'MBJ_Cross_Functions',
        'ITM_O_SL_RUN_SCRIPTING',
        'ITM_O_SL_RUN_WEB_APPL',
        'ITM_O_SL_RUN_NETSEC',
        'ITM_O_ITI_PPM_TPL',
        'ITM_O_SL_ITI_MIDDLEWARE',
        'ITM_O_SL_ITI_SERVER_WIN',
        'ITM_O_ITI_CHANGE_IMPL_DB',
        'ITM_O_SL_RUN_SAP_BASIS',
        'ITM_O_TL_CoC_FS_SO',
        'ITM_O_ITI_CHANGE_CO_WIN',
        'ITM_O_RUN_CHANGE_IMPL_BACKUP',
        'ITM_O_SL_ITI_STORAGE',
        'ITM_O_ITI_CHANGE_IMPL_WIN',
        'ITM_O_ITI_CHANGE_MGR',
        'ITM_O_SL_RUN_SERVER_WIN',
        'ITM_O_RUN_CHANGE_IMPL_VM',
        'EDC_STR_MW_DATATRANSPORT',
        'ITM_O_SL_ITI_WEB_APPL',
        'ITM_O_ITI_CHANGE_IMPL_BKUP',
        'ITM_O_RUN_CHANGE_IMPL_MON',
      ]
      security_sg = [
        'MBJ_Security_MBJ',
        'MBJ_Security_RDJ',
        'ITM_O_SL_ITI_SECURITY',
        'ITM_O_SL_RUN_SECURITY',
        'GOA_ESMO_L2',
        'CDCSP_SCCM_L2',
      ]

      if (service_group in nw_sg):
        return "Network"
      elif (service_group in server_sg):
        return "Server"
      elif (service_group in security_sg):
        return "Security"
      elif (service_group in oa_sg):
        return "OA"
      elif (service_group in swt_sg):
        return "SWT"
      elif (service_group in crm_sg):
        return "CRM"
      elif (service_group in daw_sg):
        return "DAW"
      elif (service_group in dp_sg):
        return 'Dealer Portal'
      elif (service_group in dsd_sg and 'MMCR' in app_desc):
        return 'MmcR'
      elif (service_group in pg_sg):
        return 'Paragon'
      else:
        return 'OTHER'

  def exclude_fs(self, sg):
    fs_sg = [
      'ITM_O_BU_DFS_BLAZESUPPORT',
      'ITM_O_SL_CoC_SWT_FICO_FS',
      'ITM_O_SL_DFS_SAP_AUTH',
      'ITM_O_SL_SDC_FS_SO',
      'ITM_O_TL_CoC_FS_NETSOL',
      'ITM_O_TL_CoC_FS_PEGA',
      'ITM_O_TL_CoC_FS_REPORTING',
      'ITM_O_TL_CoC_SWT_FICO_FS',
      'ITM_O_TL_DIG_FS_SO',
      'ITM_O_TL_SDC_SWT_FICO_FS',
      'MBJ_SL_AMS_BAMIC',
      'MBJ_SL_AMS_CMS',
      'MBJ_SL_AMS_COS',
      'MBJ_SL_AMS_CSS',
      'MBJ_SL_AMS_DH',
      'MBJ_SL_AMS_FSCALLCENTER',
      'MBJ_SL_AMS_FSSAP',
      'MBJ_SL_AMS_JPOS',
      'MBJ_SL_AMS_LIVELINK',
      'MBJ_SL_AMS_PEGA',
      'MBJ_SL_AMS_RNS',
      'MBJ_SL_AMS_WFS',
      'MBJ_SL_FS_ACC',
      'MBJ_SL_FS_BAMIC',
      'MBJ_SL_FS_CMS',
      'MBJ_SL_FS_Concur',
      'MBJ_SL_FS_COS',
      'MBJ_SL_FS_CSS',
      'MBJ_SL_FS_DMS',
      'MBJ_SL_FS_DX',
      'MBJ_SL_FS_Finance',
      'MBJ_SL_FS_FSCALLCENTER',
      'MBJ_SL_FS_FSSAP',
      'MBJ_SL_FS_JPOS',
      'MBJ_SL_FS_Leasing1',
      'MBJ_SL_FS_LIVELINK',
      'MBJ_SL_FS_PEGA',
      'MBJ_SL_FS_RNS',
      'MBJ_SL_FS_WFS',
    ]
    if (sg in fs_sg):
      return True
    else:
      return False

  def create_category(self, category, classification):
    if category == "Incident":
      if classification == "INQUIRY":
        return 'Inquiry'
      else:
        return 'Problem'
    else:
      if classification == "INQUIRY":
        return "Inquiry"
      elif classification == "REQUEST FOR CHANGE":
        return "Change"
      else:
        problem_class = [
          "PROBLEM_ADMINISTRATION/CONFIGURATION",
          "PROBLEM_COMMUNICATION",
          "PROBLEM_DEFECT/DESTRUCTION",
          "PROBLEM_QUALITY OF SERVICE/PERFORMANCE",
          "PROBLEM_SYNCHRONISATION",
          "PROBLEM_USER HANDLING",
        ]
        id_class = [
          'ACCOUNTING',
          'ADMINISTRATION_AUTHORISATIONS',
          'ADMINISTRATION_CHANGE',
          'ADMINISTRATION_COMBINATION',
          'ADMINISTRATION_CREATE',
          'ADMINISTRATION_DELETE',
          'ADMINISTRATION_LOCK',
          'ADMINISTRATION_LOGIN(RESET PASSWORD)',
          'ADMINISTRATION_RESOURCE',
          'ADMINISTRATION_UNLOCK',
          'ADMINISTRATION_USER',
          'SPECIAL_COC OVERSEAS_ACCESS_ADMINISTRATION CHANGE',
          'SPECIAL_COC OVERSEAS_ACCESS_ADMINISTRATION CREATE',
          'SPECIAL_COC OVERSEAS_ACCESS_ADMINISTRATION DELETE',
          'SPECIAL_COC OVERSEAS_ACCESS_ADMINISTRATION LOGIN(',
          'SPECIAL_COC OVERSEAS_PRD_ACCESS_ADMINISTRATION CHA',
          'SPECIAL_COC OVERSEAS_PRD_ACCESS_ADMINISTRATION CRE',
          'SPECIAL_COC OVERSEAS_PRD_ACCESS_ADMINISTRATION DEL',
          'SPECIAL_COC OVERSEAS_PRD_ACCESS_ADMINISTRATION LOG',
          'SPECIAL_COC OVERSEAS_PRD_REQUEST_ACCESS',
        ]

        if (classification in problem_class):
          return "Problem"
        elif (classification in id_class):
          return "ID"
        else:
          return "Other"

  def readExcel(self, isOpen):
    # Open Excel
    wb = xw.Book(self.import_file)
    sheet = wb.sheets(1)

    cism = {}
    start_row = 2
    last_row = sheet.range('A1').current_region.last_cell.row
    # last_row = 100

    for i in range(start_row, last_row + 1):
      try:
        service_group = sheet.range(f'Y{i}').value
        # Exclude MBF tickets
        if (self.exclude_fs(service_group)):
          continue

        # Tiket ID
        id = sheet.range(f'A{i}').value

        # アプリ担当チームを抽出 => [CRM/Indcident] => CRM
        team = self.find_team(
          sheet.range(f'Q{i}').value, service_group)

        # アプリ担当チーム
        cism[id] = {"team": team}

        # Request or Incident
        ticket_type = sheet.range(f'N{i}').value
        category = 'Request' if ticket_type == 'Request' else "Incident"
        cism[id]["category"] = category

        # どこから問い合わせがあったか
        cism[id]["source"] = sheet.range(f'J{i}').value

        # status
        status = sheet.range(f'O{i}').value
        cism[id]["status"] = status

        # Classification => Category
        classification = sheet.range(f'S{i}').value
        # cism[id]["classification"] = classification
        cism[id]["category_1"] = self.create_category(
          category, classification)

        # TODO 本来はここで１８ケ月前のものははじくべき
        # チケット起票された月
        create_date = sheet.range(f'D{i}').value
        create_month = str(create_date.year) + \
          str(create_date.month).zfill(2)
        cism[id]["create_date"] = create_month

        # Closeまでにかかった日付
        ticket_age = sheet.range(f'BQ{i}').value
        cism[id]["ticket_age"] = ticket_age

        # Open ticket のエクセルか判断させる為
        cism[id]["isOpen"] = isOpen

        # closeした月 yyyymm (201909)
        # 現時点でCloseになっていないとブランクになっている
        # if status == "Closed":
        #   close_date = sheet.range(f'F{i}').value
        #   close_month = str(close_date.year) + \
        #         str(close_date.month).zfill(2)
        #   cism[id]["close_date"] = close_month
        # else:
        #   print("Start date: ", create_date, create_month)

        # pprint.pprint(cism[id])

      except Exception as e:
        print(f'Error parsing excel => ticket# : {id}')
        print(e)
        import traceback
        traceback.print_exc()

    # Close Excel
    # wb.save(self.import_file)
    wb.app.quit()

    return cism


class AnalyzeCism:
  def __init__(self, all_cism):
    self.all_cism = all_cism

  def count(self, ticket_data):
    for v in self.all_cism.values():
      try:
        app_team = v['team']

        # Open ticket のエクセル場合は、ロングオープンの分析
        if not v['isOpen']:
          ticket_month = ticket_data[v['create_date']]

          if v['category'] == "Incident":
            ticket_month["Incident"]["sum"] += 1

            if v["category_1"] == 'Inquiry':
              ticket_month["Incident"]["inquiry"]["sum"] += 1
              ticket_month["Incident"]["inquiry"][app_team] += 1
            else:
              ticket_month["Incident"]["problem"]["sum"] += 1
              ticket_month["Incident"]["problem"][app_team] += 1
          else:
            ticket_month["Request"]["sum"] += 1

            if v['category_1'] == "ID":
              ticket_month['Request']["id"]["sum"] += 1
              ticket_month['Request']["id"][app_team] += 1

              if v['source'] == "SAPI" and int(v['create_date']) >= 201901:
                ticket_month['Request']["id"]["daw"][app_team] += 1
              elif app_team == 'DAW':
                ticket_month['Request']["id"]["daw"][app_team] += 1
              else:
                ticket_month['Request']["id"]["not_daw"][app_team] += 1
            else:
              ticket_month['Request']["other"]["sum"] += 1
              ticket_month['Request']["other"][app_team] += 1

          ticket_month["total"][app_team] += 1

        else:
          # Long Open 分析はこっち
          if v['status'] == "Solved":
            continue

          if v['ticket_age'] > 20:
            ticket_data[last_month]["longopen"]["sum"] += 1
            ticket_data[last_month]["longopen"][app_team] += 1

          ticket_data[last_month]["open"]["sum"] += 1
          ticket_data[last_month]["open"][app_team] += 1

      except Exception as e:
        print("Error duing counting ... ")
        print(e)
        print(v)
        import traceback
        traceback.print_exc()


class SoReport:
  def __init__(self):
    self.ticket_data = {}
    self.output_dir = os.path.join(cwd, 'output', last_month)
    self.output_file = os.path.join(self.output_dir, 'output.xlsx')
    self.output_longopen_dir = os.path.join(cwd, 'output', 'longopen')
    self.output_longopen_file = os.path.join(
        self.output_longopen_dir, 'longopen_history.xlsx')

  def create_key(self, ticket_data):
    today = datetime.date.today()
    # 18ヶ月前からi1ヶ月前までトラッキングする
    for i in reversed(range(1, track_month+1)):
      month = today - relativedelta(months=i)
      ticket_date_key = month.strftime("%Y%m")
      ticket_data[ticket_date_key] = {}

  def create_ticket_data(self):
    self.create_key(self.ticket_data)
    for k in self.ticket_data.keys():
      self.ticket_data[k] = {
        "Incident": {"sum": 0,
               "inquiry": {"sum": 0},
               "problem": {"sum": 0},
               },
        "Request": {"sum": 0,
              "id": {"sum": 0, "daw": {}, "not_daw":{}},
              "other": {"sum": 0},
              },
        "open": {"sum": 0},
        "longopen": {"sum": 0},
        "total": {}
      }

      for app in apps:
        self.ticket_data[k]["total"][app] = 0
        self.ticket_data[k]["Incident"]["inquiry"][app] = 0
        self.ticket_data[k]["Incident"]["problem"][app] = 0
        self.ticket_data[k]["Request"]["id"][app] = 0
        self.ticket_data[k]["Request"]["id"]["daw"][app] = 0
        self.ticket_data[k]["Request"]["id"]["not_daw"][app] = 0
        self.ticket_data[k]["Request"]["other"][app] = 0
        self.ticket_data[k]["open"][app] = 0
        self.ticket_data[k]["longopen"][app] = 0

    return self.ticket_data

  def dump_result_to_excel(self):
    # Open Excel (新規)
    wb2 = xw.Book()

    # Sheet1 にoverviewを保存
    self.dump_overview(sheet=wb2.sheets(1))

    # アプリ毎の詳細の
    titles = ["total",
              "Incident_inquiry", "Incident_problem",
              "Request_id", "Request_other",
              "ID_daw", "ID_not_daw"]
    prev_sheet='Sheet1'

    for title in titles:
      sheet = wb2.sheets.add(name=title, after=prev_sheet)
      prev_sheet = title

      if title == "total":
        self.dump_ticket(sheet=sheet, category="total")
      elif title == "Incident_inquiry":
        self.dump_ticket(sheet=sheet, category="Incident", category1="inquiry")
      elif title == "Incident_problem":
        self.dump_ticket(sheet=sheet, category="Incident", category1="problem")
      elif title == "Request_id":
        self.dump_ticket(sheet=sheet, category="Request", category1="id")
      elif title == "Request_other":
        self.dump_ticket(sheet=sheet, category="Request", category1="other")
      elif title == "ID_daw":
        self.dump_ticket(sheet=sheet, category="Request", category1="id", source="daw")
      elif title == "ID_not_daw":
        self.dump_ticket(sheet=sheet, category="Request", category1="id", source="not_daw")

    # Save & Close Excel
    wb2.save(self.output_file)
    wb2.app.quit()
  
  def dump_overview(self, sheet):
    # application team (Other以外) = [[app1],[app2]..]
    arr = [
      [''], ["Incident"], ["Request"], ["total"],
      [''],
      [''], ["inquiry"], ["problem"], ["total"],
      [''],
      [''], ["id"], ["other"], ["total"]   
      ]

    # application team (Other以外) を縦に登録 & 幅13に設定
    sheet.range('B2').value = arr
    sheet.range('B2').columns.rng.column_width = 13

    # 年月とアプリ毎のチケット数を格納する配列
    list_ticket = [[] for i in range(len(arr))]

    # 年月とチケット数を配列に格納
    for k, v in self.ticket_data.items():
      list_ticket[0].append(k)
      list_ticket[1].append(v['Incident']['sum'])
      list_ticket[2].append(v['Request']['sum'])
      list_ticket[3].append(v['Incident']['sum'] + v['Request']['sum'])

      list_ticket[5].append(k)
      list_ticket[6].append(v['Incident']['inquiry']['sum'])
      list_ticket[7].append(v['Incident']['problem']['sum'])
      list_ticket[8].append(v['Incident']['sum'])
      
      list_ticket[10].append(k)
      list_ticket[11].append(v['Request']["id"]['sum'])
      list_ticket[12].append(v['Request']["other"]['sum'])
      list_ticket[13].append(v['Request']['sum'])

    # 年月とアプリごとのチケット数を埋める（左から右）
    # 2 は行の始まり TODO ハードコーディングやめる
    for i, item in enumerate(list_ticket):
      sheet.range(f'c{2+i}').value = item

  def update_longopen_to_excel(self):
    longopen_data = self.ticket_data[last_month]['longopen']
    wb3 = xw.Book(self.output_longopen_file)
    
    sheet = wb3.sheets['overview']
    # 配列の右を取得
    last_right_column = sheet.range('C2').end('right')
    if int(last_right_column.value) == int(last_month):
      cur_position = last_right_column
    else:
      cur_position = last_right_column.offset(0, 1)

    # 先月を入れる
    cur_position.value = last_month
    cur_position = cur_position.offset(1, 0)
    cur_position.value = self.ticket_data[last_month]['open']['sum'] - \
        longopen_data['sum']
    cur_position = cur_position.offset(1, 0)
    cur_position.value = longopen_data['sum']
    cur_position = cur_position.offset(1, 0)
    cur_position.value = self.ticket_data[last_month]['open']['sum']
    
    
    sheet = wb3.sheets['breakdown']
    # 配列の右を取得
    last_right_column = sheet.range('C2').end('right')
    if int(last_right_column.value) == int(last_month):
      cur_position = last_right_column
    else:
      cur_position = last_right_column.offset(0, 1)

    # 先月を入れる
    cur_position.value = last_month
    cur_position = cur_position.offset(1, 0)

    # app_list = []
    app_list_start = 'B3'
    app_list_last = sheet.range(app_list_start).end('down')
    app_list = sheet.range(app_list_start, app_list_last).value

    # 左列のアプリ名を取得し、それに対するチケット数を追加
    for i, app_item in enumerate(app_list):
      cur_position.offset(i, 0).value = longopen_data[app_item]

    # Sort
    self.sort_descend(sheet)

    # Save & Close
    wb3.save(self.output_longopen_file)
    # wb3.app.quit()

  def dump_ticket(self, **kwargs):
    sheet = kwargs['sheet']
    category = kwargs['category']
    category1 = kwargs['category1'] if len(kwargs) >= 3 else ''
    source = kwargs['source'] if len(kwargs) == 4 else ''

    # application team (Other以外) = [[app1],[app2]..]
    arr_app = []
    for app in apps:
      if app == "OTHER":
        continue
      arr_app.append([app])

    # application team (Other以外) を縦に登録
    sheet.range('B3').value = arr_app
    sheet.range('B3').columns.rng.column_width = 13

    # 年月とアプリ毎のチケット数を格納する配列
    list_ticket = [[] for i in range(len(arr_app) + 1)]

    for k, v in self.ticket_data.items():
      for i in range(len(arr_app) + 1):
        if i == 0:
          list_ticket[i].append(k)
          continue

        if len(kwargs) == 2:
          list_ticket[i].append(v[category][arr_app[i - 1][0]])
        elif len(kwargs) == 3:
          list_ticket[i].append(v[category][category1][arr_app[i - 1][0]])
        else:
          list_ticket[i].append(v[category][category1][source][arr_app[i - 1][0]])

    # 年月とアプリごとのチケット数を埋める（左から右）
    # 2 は行の始まり TODO ハードコーディングやめる
    for i, item in enumerate(list_ticket):
      sheet.range(f'c{2+i}').value = item

    # Sort
    self.sort_descend(sheet)

  def sort_descend(self, sheet):
    left_top_cell = 'B3'
    right_top_cell = sheet.range(left_top_cell).end('right')
    right_bottom_cell = right_top_cell.end('down')

    sheet.range(left_top_cell, right_bottom_cell).api.Sort(
        Key1=sheet.range(right_top_cell, right_bottom_cell).api, Order1=2)

  def draw_ticket_overview(self):
    arr_month = []
    arr_request = []
    arr_incident = []

    for k, v in self.ticket_data.items():
      arr_month.append(k)
      arr_request.append(v["Incident"]["sum"])
      arr_incident.append(v["Request"]["sum"])

    left = np.array(arr_month)
    height1 = np.array(arr_request)
    height2 = np.array(arr_incident)
    p1 = plt.bar(left, height1, color="green")
    p2 = plt.bar(left, height2, bottom=height1, color="orange")
    plt.legend((p1[0], p2[0]), ('Request', 'Incident'))
    plt.show()

if __name__ == '__main__':
  start = time.time()

  if not os.path.exists(os.path.join(cwd, 'import', last_month)):
    print('*************************************************')
    print(f'Please create directory {last_month} in import')
    sys.exit()
  # Pasrse Excel
  # TODO loop にするべきでは？

  analyze2018_1 = AnalyzeExcel("AK_Ticket_2018-1.xlsx")
  all_cism.update(analyze2018_1.readExcel(isOpen=False))

  analyze2018_2 = AnalyzeExcel("AK_Ticket_2018-2.xlsx")
  all_cism.update(analyze2018_2.readExcel(isOpen=False))

  analyze2019_1 = AnalyzeExcel("AK_Ticket_2019-1.xlsx")
  all_cism.update(analyze2019_1.readExcel(isOpen=False))

  analyze2019_2 = AnalyzeExcel("AK_Ticket_2019-2.xlsx")
  all_cism.update(analyze2019_2.readExcel(isOpen=False))

  analyzeTicketOpen = AnalyzeExcel("AK_Ticket_Open.xlsx")
  all_cism.update(analyzeTicketOpen.readExcel(isOpen=True))

  # crating ticket_data
  soReport = SoReport()
  ticket_data = soReport.create_ticket_data()

  # Analyze all ticket data for so report
  analyzeCism = AnalyzeCism(all_cism)
  analyzeCism.count(ticket_data)

  # Debug 用
  pprint.pprint(ticket_data)

  # Dump
  soReport.dump_result_to_excel()
  soReport.update_longopen_to_excel()
  # soReport.draw_ticket_overview()


  end = time.time()
  elapsed_time = end - start
  print("{0.hours:02}:{0.minutes:02}:{0.seconds:02}".format(
    relativedelta(seconds=round(elapsed_time))))
