import csv

from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill

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
        columnValue = columnValue if columnValue is not None else 'N/A'
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


def create_csv(file_name, columns, dataRows):
  columnNames = []
  for columnName, columnIndex in columns:
    columnNames.append(columnName)
  try:
    with open(file_name, mode='w', newline='') as csv_out:
        writer = csv.DictWriter(csv_out, fieldnames=columnNames)
        writer.writeheader()
        writer.writerows(dataRows)
    file_name_csv = file_name + '.csv'
    print(f"Summary CSV report saved to '{file_name_csv}'.  It has {len(dataRows)} rows.")
    print(f"Summary report saved at: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
  except Exception as err:
      print(f"Error saving worksheet\n  {err}")
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
  # ws["A1"].fill = PatternFill("solid", start_color="c9c9c9")

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
  createNewMemberColumnNames = ['Last Name', 'First Name', 'Email', 'Username']
  upgradeMemberColumnNames = ['Last Name', 'First Name', 'Email', 'Website ID', 'Username']
  verifyColumnNames = ['Last Name', 'First Name', 'Email']
  
  wb = Workbook()
  ws = wb.active
  ws.title = 'Create'
  ws.append(createNewMemberColumnNames)
  for name, status in statusList:
    if status == 'non-member':
      lastName, firstName = name
      emailValue = 'N/A'
      # Do a full 'table' scan for each individual 
      for item in championRows:
        if (item[0] == firstName) and (item[1] == lastName):
          emailValue = item[2]
      if len(lastName) < 9:
        userName = lastName + firstName[0:1]
      else:
        userName = lastName[0:9] + firstName[0:1]
      userName = userName.lower()
      row = name
      row = (*row, emailValue, userName,)
      ws.append(row)
    else:
      continue
  ws.freeze_panes = 'A2'
  ws.auto_filter.ref = ws.dimensions
  # ws["A1"].fill = PatternFill("solid", start_color="c9c9c9")

  wb.create_sheet('Upgrade')
  ws = wb['Upgrade']
  ws.append(upgradeMemberColumnNames)
  for name, status in statusList:
    lastName, firstName = name
    if status == 'member':
      emailValue = 'N/A'
      for item in memberRows:
        if (item[4] == firstName) and (item[3] == lastName):
          # Column L (Website ID (RO)) in the member export spreadsheet
          memberID = item[11] ;
          # Column W in the member export spreadsheet
          memberUsername = item[22] ;
          # Column Y in the member export spreadsheet
          memberEmailValue = item[24] ;
      row = name
      row = (*row, memberEmailValue, memberID, memberUsername, )
      ws.append(row)
    else:
      continue
  ws.freeze_panes = 'A2'
  ws.auto_filter.ref = ws.dimensions
  # ws["A1"].fill = PatternFill("solid", start_color="c9c9c9")

  wb.create_sheet('Verify')
  ws = wb['Verify']
  ws.append(verifyColumnNames)
  for name, status in statusList:
    if status == 'neither':
      lastName, firstName = name
      emailValue = 'N/A'
        # Do a full 'table' scan for each individual 
      for item in championRows:
        if (item[0] == firstName) and (item[1] == lastName):
          emailValue = item[2]
      row = name
      row = (*row, emailValue, )
      ws.append(row)
    else:
      continue
  ws.freeze_panes = 'A2'
  ws.auto_filter.ref = ws.dimensions
  # ws["A1"].fill = PatternFill("solid", start_color="c9c9c9")

  try:
    file_name_xlsx = file_name + '.xlsx'
    wb.save(file_name_xlsx)
    print(f"Summary Excel report saved to '{file_name_xlsx}.xlsx'.")

  except PermissionError:
      print(f"Could not save worksheet: '{file_name_xlsx}'\nCheck if a previous version is open in Excel.")
      # sg.popup_error("Could not save 'account_summary' report\nCheck if a previous version is open in Excel.")
      # window['-STATUS-'].update("'Save Report' operation canceled.")
  except Exception as err:
      print(f"Error saving worksheet\n  {err}")
      # sg.popup_error(f"Error saving account_summary report\n  {err}")
  finally:
      wb.close()


def find_champion_status(championNames, memberNames, nonMemberNames):
  statusList = []
  for champion in championNames:
    if champion in memberNames:
      status = 'member'
    elif champion in nonMemberNames:
      status = 'non-member'
    else:
      status = 'neither'
    statusList.append((champion, status,))

  histogramCount = {}
  for champion, status in statusList:
    try:
        histogramCount[status] += 1
    except KeyError:
        histogramCount[status] = 1
  print(f"\nChampion status histogram = {histogramCount}")

  return statusList


# bvb TODO Add relative Path to the data folder
if __name__ == "__main__":

  memberRows = xlsx_reader("pqi-etl/data/20221125_members_export_excel_ymx.xlsx")
  # Exclude header row from row count
  print(f"Members (xlsx): {len(memberRows) - 1}")
  columns = [('Last Name', 3,), ('First Name', 4)]
  memberNames = create_name_dataset(memberRows, columns)
  print(f"Member unique names: {len(memberNames)}")
  create_xlsx('pqi-etl/data/member_unique_names', 'names', columns, memberNames)

  nonMemberRows = xlsx_reader("pqi-etl/data/20221125_non_members_export_excel_ymx.xlsx")
  columns = [('Last Name', 2,), ('First Name', 3)]
  # Exclude header row from row count
  print(f"\nNon-Members (xlsx): {len(nonMemberRows) - 1}")
  nonMemberNames = create_name_dataset(nonMemberRows, columns)
  print(f"Non-Member unique names: {len(nonMemberNames)}")
  create_xlsx('pqi-etl/data/non_member_unique_names', 'names', columns, nonMemberNames)

  championRows = xlsx_reader("pqi-etl/data/Master List of all SPEAK UP Champion Attendees.xlsx")
  columns = [('Last Name', 1,), ('First Name', 0)]
  # columns = [('Last Name', 1,), ('First Name', 0), ('Email', 2)]
  # Exclude header row from row count
  print(f"\nChampions (xlsx): {len(championRows) - 1}")
  championNames = create_name_dataset(championRows, columns)
  print(f"Champion unique names: {len(championNames)}")
  create_xlsx('pqi-etl/data/champion_unique_names', 'names', columns, championNames)

  fileName = "pqi-etl/data/champion_email_action.xlsx"
  statusList = find_champion_status(championNames, memberNames, nonMemberNames)
  create_champion_actions_xlsx(fileName, statusList, memberRows, nonMemberRows, championRows)
