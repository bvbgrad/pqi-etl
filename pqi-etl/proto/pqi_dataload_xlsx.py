import csv
import secrets

from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill

RI = {'A':0, 'B':1, 'C':2, 'D':3, 'E':4, 'F':5, 'G':6, 'H':7, 'I':8, 'J':9, 
  'K':10, 'L':11, 'M':12, 'N':13, 'O':14, 'P':15, 'Q':16, 'R':17, 'S':18,
  'T':19, 'U':20, 'V':21, 'W':22, 'X':23, 'Y':24, 'Z':25}
NewMemberCols = {
  "Member Type Code": "",
  "Username": "",
  "Password": "",
  "First Name": "",
  "Last Name": "",
  "Email Address": "",
  "Middle Name": "",
  "Maiden Name": "",
  "Nickname": "",
  "Member Name Title": "Professional Title",
  "Member Name Suffix": "Name Suffix",
  "Gender": "",
  "Registration Date": "",
  "Member Approved": "",
  "Membership": "",
  # "Date Membership Expires": "",
  # "Membership Expires": "",
  "Date Last Renewed": "",
  "Email Bounced": "",
  "Home Address Line 1": "",
  "Home Address Line 2": "",
  "Home City": "",
  "Home Location": "",
  "Home Postal Code": "",
  "Home Country": "",
  "Home Phone Area Code": "",
  "Home Phone": "",
  "Mobile Area Code": "",
  "Mobile": "",
  "Employer Name": "Organization Name",
  "Professional Title": "",
  "Profession": "",
  "Employer Address Line 1": "",
  "Employer Address Line 2": "",
  "Employer City": "",
  "Employer Location": "",
  "Employer Postal Code": "",
  "Employer Country": "",
  "Employer Phone Area Code": "",
  "Employer Phone": "",
  "Employer Website": "",
  "Internal Comments": "",
  "Champion": ""
}

def generate_username(keyValues):
  lastName, firstName, emailValue = keyValues
  if len(lastName) < 9:
    userName = lastName + firstName[0:1]
  else:
    userName = lastName[0:9] + firstName[0:1]
  userName = userName.lower()
  return userName


def generate_password():
  return secrets.token_urlsafe(8)

def find_download_data_row(dataRows, eMail, emailColNum):
  for dataRow in dataRows:
    if dataRow[emailColNum] == eMail:
      return dataRow

def get_data_value(row, indexNum):
  if row is not None:
    rowList = list(row)
    value = rowList[indexNum]
    if value is None:
      value = ''
  else:
    print(f"Index: {indexNum} row {row}")
  return value

def fill_row_values(keyValues, championRow, nonMemberRow):
  memberTypeCode = 'BasicI'
  username = generate_username(keyValues)
  password = generate_password()
  middleName = nonMemberRow[RI['F']]
  nameSuffix = nonMemberRow[RI['H']]
  organization = championRow[RI['D']]
  websiteId = nonMemberRow[RI['K']]

  # middleName = get_data_value(nonMemberRow, RI['F'])
  # nameSuffix = get_data_value(nonMemberRow, RI['H'])
  # organization = get_data_value(championRow, RI['D'])
  # websiteId = get_data_value(nonMemberRow, RI['K'])

  row = ( memberTypeCode, *keyValues, websiteId, username, password, middleName, nameSuffix, organization,)

  return row


def xlsx_reader(filename):
    rows = []
    try:
        wb = load_workbook(filename=filename, read_only=True)
        print(f"Worksheets: {wb.sheetnames}")
        ws = wb[wb.sheetnames[0]]
        for row in ws.values:
            rows.append(row)
        wb.close()
    except Exception as e:
        print(f"Supported formats are: .xlsx,.xlsm,.xltx,.xltm \n  {e}")
    wb.close()
    return rows

def create_name_dataset(rows, columns):
  lookupSet = set()
  lookupList = []
  print(f"Extracting data for columns: {columns}")
  for i, row in enumerate(rows, 1):
    # Throw away the header row
    if i == 1: 
      continue
    rowTuple = ()
    try:
      for columnName, columnIndex in columns:
        columnValue = row[columnIndex]
        columnValue = columnValue if columnValue is not None else ''
        rowTuple = (*rowTuple, columnValue,)
      lookupSet.add(rowTuple)
      lookupList.append(rowTuple)
    except Exception as e:
      print(f"Data error on row {i:5}. {e}")
      print(f"  {row}")
  sortedDataList = sorted(lookupList)
  print(f"Sorted data list: {len(sortedDataList)}")
  sortedDataset = sorted(lookupSet)
  print(f"Sorted data set: {len(sortedDataset)}")
  
  return sortedDataset

def create_csv(file_name_csv, dataRows):
  try:
    with open(file_name_csv, mode='w', newline='') as csv_out:
        writer = csv.writer(csv_out)
        writer.writerows(dataRows)
    print(f"Summary CSV report saved to '{file_name_csv}'.  It has {len(dataRows)} rows.")
    print(f"Summary report saved at: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
  except Exception as err:
      print(f"Error saving csv worksheet\n  {err}")
      # sg.popup_error(f"Error saving account_summary report\n  {err}")


def create_xlsx(file_name, sheetName, columns, dataRows):
  columnNames = []
  for columnName, columnIndex in columns:
    columnNames.append(columnName)
  wb = Workbook()
  ws = wb.active
  ws.title = sheetName
  ws.append(columnNames)
  for row in dataRows:
      ws.append(row)
  ws.freeze_panes = 'A2'
  ws.auto_filter.ref = ws.dimensions
  ws["A1"].fill = PatternFill("solid", start_color="c9c9c9")

  try:
    file_name_xlsx = file_name + '.xlsx'
    wb.save(file_name_xlsx)
    print(f"Summary Excel report saved to '{file_name_xlsx}.xlsx'.  It has {len(dataRows)} rows.")

  except PermissionError:
      print(f"Could not save worksheet: '{file_name_xlsx}'\nCheck if a previous version is open in Excel.")
      # sg.popup_error("Could not save 'account_summary' report\nCheck if a previous version is open in Excel.")
      # window['-STATUS-'].update("'Save Report' operation canceled.")
  except Exception as err:
      print(f"Error saving worksheet\n  {err}")
      # sg.popup_error(f"Error saving account_summary report\n  {err}")
  finally:
      wb.close()


def create_champion_actions_xlsx(file_name, statusList, memberRows, nonMemberRows, championRows):
  createNewMemberColumnNames = ['Last Name', 'First Name', 'Email', 'Username', 'Password']
  upgradeMemberColumnNames = ['Last Name', 'First Name', 'Email', 'Username', 'Website ID', 'Champion']
  verifyColumnNames = ['Last Name', 'First Name', 'Email']
  
  wb = Workbook()
  ws = wb.active
  ws.title = 'Create'
  ws.append(createNewMemberColumnNames)
  for keyValues, status in statusList:
    if status == 'non-member':
      lastName, firstName, emailValue = keyValues
      championRow = find_download_data_row(championRows, emailValue, RI['C'])
      nonMemberRow = find_download_data_row(nonMemberRows, emailValue, RI['U'])
      if championRow is None or nonMemberRow is None:
        print(f"Skipping {keyValues}")
      else:
        row = fill_row_values(keyValues, championRow, nonMemberRow)
        ws.append(row)
    else:
      continue
  ws.freeze_panes = 'A2'
  ws.auto_filter.ref = ws.dimensions
  ws["A1"].fill = PatternFill("solid", start_color="c9c9c9")

  wb.create_sheet('Upgrade')
  ws = wb['Upgrade']
  ws.append(upgradeMemberColumnNames)
  for keyValues, status in statusList:
    lastName, firstName, emailValue = keyValues
    if status == 'member':
      for item in memberRows:
  # Column Y (zero based index = 24) contains the email value in the member export spreadsheet
        if item[24] == emailValue:
          # Column L (Website ID (RO)) in the member export spreadsheet
          memberID = item[11] ;
          # Column W in the member export spreadsheet
          memberUsername = item[22] ;
          championFlag = 'Yes' ;
          row = keyValues
          row = (*row, memberUsername, memberID, championFlag)
          ws.append(row)
    else:
      continue
  ws.freeze_panes = 'A2'
  ws.auto_filter.ref = ws.dimensions
  ws["A1"].fill = PatternFill("solid", start_color="c9c9c9")

  wb.create_sheet('Verify')
  ws = wb['Verify']
  ws.append(verifyColumnNames)
  for keyValues, status in statusList:
    if status == 'verify':
      ws.append(row)
    else:
      continue
  ws.freeze_panes = 'A2'
  ws.auto_filter.ref = ws.dimensions
  ws["A1"].fill = PatternFill("solid", start_color="c9c9c9")

  try:
    wb.save(file_name)
    print(f"Summary Excel report saved to '{file_name}'.")

  except PermissionError:
      print(f"Could not save worksheet: '{file_name}'\nCheck if a previous version is open in Excel.")
      # sg.popup_error("Could not save 'account_summary' report\nCheck if a previous version is open in Excel.")
      # window['-STATUS-'].update("'Save Report' operation canceled.")
  except Exception as err:
      print(f"Error saving worksheet\n  {err}")
      # sg.popup_error(f"Error saving account_summary report\n  {err}")
  finally:
      wb.close()


def find_champion_status(championNames, memberNames, nonMemberNames):
  statusList = []
  count = 0
  for champion in championNames:
    status = 'verify'
    lastNameChampion, firstNameChampion, emailChampion = champion
    emailChampion = emailChampion.lower()
    for keyValues in memberNames:
      lastName, firstName, emailMember = keyValues
      if emailChampion == emailMember.lower():
        status = 'member'
    for keyValues in nonMemberNames:
      lastName, firstName, emailNonMember = keyValues
      if emailChampion == emailNonMember.lower():
        status = 'non-member'
    if emailChampion is None or emailChampion == '':
      status = 'verify'
  # Email check failed: see if can find in YM downloads using First & Last names
    if status == 'verify':
      for keyValues in memberNames:
        lastName, firstName, emailMember = keyValues
        if (firstNameChampion == firstName) and (lastNameChampion == lastName):
          if emailChampion != emailMember:
            print(f"{firstNameChampion} {lastNameChampion} status changed based on name match:")
            print(f"   Champion email = '{emailChampion}' - Member email = '{emailMember}'")
          status = 'member'
          count += 1
          print(f"   Status changed from 'verify' to '{status}'.")
          break
      for keyValues in nonMemberNames:
        lastName, firstName, emailNonMember = keyValues
        if (firstNameChampion == firstName) and (lastNameChampion == lastName):
          if emailChampion != emailNonMember:
            print(f"{firstNameChampion} {lastNameChampion} status changed based on name match:")
            print(f"   Champion email = '{emailChampion}' - Non-member email = '{emailNonMember}'")
          status = 'non-member'
          count += 1
          print(f"   Status changed from 'verify' to '{status}'.")
  # Need champion key values in next stage of the process
    statusList.append((champion, status,))

  if count == 0:
    print("No status changes based on a name match process rather than an email match")
  else:
    print(f"There were {count} status changes based on a name match process rather than an email match")

  histogramCount = {}
  for champion, status in statusList:
    try:
        histogramCount[status] += 1
    except KeyError:
        histogramCount[status] = 1
  print(f"\nChampion status histogram = {histogramCount}")
  numChampionNames = len(championNames)
  numStatus = len(statusList)
  if numChampionNames == numStatus:
    print(f"All {numChampionNames} champion records have been processed.")
  else:
    print(f"Error: {numChampionNames} champion records does not match {numStatus} status determinations.")

  return statusList


# bvb TODO Add relative Path to the data folder
if __name__ == "__main__":

  memberRows = xlsx_reader("pqi-etl/data/20221125_members_export_excel_ymx.xlsx")
  # Exclude header row from row count
  print(f"Members (xlsx): {len(memberRows) - 1}")
  # Column Y (zero based index = 24) contains the email value in the member export spreadsheet
  columns = [('Last Name', 3,), ('First Name', 4,), ('Email Address', 24,)]
  memberNames = create_name_dataset(memberRows, columns)
  print(f"Member unique names: {len(memberNames)}")
  create_xlsx('pqi-etl/data/member_unique_names', 'names', columns, memberNames)

  nonMemberRows = xlsx_reader("pqi-etl/data/20221125_non_members_export_excel_ymx.xlsx")
  # Column U (zero based index = 20) contains the email value in the non_member export spreadsheet
  columns = [('Last Name', 2,), ('First Name', 3,), ('Email Address', 20,)]
  # Exclude header row from row count
  print(f"\nNon-Members (xlsx): {len(nonMemberRows) - 1}")
  nonMemberNames = create_name_dataset(nonMemberRows, columns)
  print(f"Non-Member unique names: {len(nonMemberNames)}")
  create_xlsx('pqi-etl/data/non_member_unique_names', 'names', columns, nonMemberNames)

  championRows = xlsx_reader("pqi-etl/data/Master List of all SPEAK UP Champion Attendees.xlsx")
  # Column C (zero based index = 2) contains the email value in the champion export spreadsheet
  columns = [('Last Name', 1,), ('First Name', 0), ('Email', 2)]
  # Exclude header row from row count
  print(f"\nChampions (xlsx): {len(championRows) - 1}")
  championNames = create_name_dataset(championRows, columns)
  print(f"Champion unique names: {len(championNames)}")
  create_xlsx('pqi-etl/data/champion_unique_names', 'names', columns, championNames)

  fileName = "pqi-etl/data/champion_email_action.xlsx"
  statusList = find_champion_status(championNames, memberNames, nonMemberNames)
  create_champion_actions_xlsx(fileName, statusList, memberRows, nonMemberRows, championRows)

  fileNameCsv = "pqi-etl/data/champion_update_action.csv"
  upgradeCsvColumnNames = ['Website ID', 'Champion']
  dataRows = []
  dataRows.append(upgradeCsvColumnNames)
  for keyValues, status in statusList:
    lastName, firstName, emailValue = keyValues
    if status == 'member':
      for i, item in enumerate(memberRows, 1):
  # Column Y (email address)) in the member export spreadsheet
        if item[24] is not None:
          emailMember = item[24].lower()
          if emailMember == emailValue.lower():
    # Column L (Website ID (RO)) in the member export spreadsheet
            memberID = item[11] ;
            championFlag = 'Yes' ;
            row = list((memberID, championFlag,))
            dataRows.append(row)
  create_csv(fileNameCsv, dataRows)
