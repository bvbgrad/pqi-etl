import sys
import codecs
import os.path
import io
import argparse
from pandas.core.algorithms import duplicated
from tabulate import tabulate
import pandas as pd

import secrets
import string

from datetime import datetime

ArgV = argparse.ArgumentParser(description='Output list of champtions to import')
ArgV.add_argument('--inchampions', dest='InputChampionsFile', required=True, help='Excel master champions list')
ArgV.add_argument('--inmembers', dest='InputMembersFile', required=True, help='Exported member database file (UTF8 Batch Export')
ArgV.add_argument('--innonmembers', dest='InputNonMembersFile', required=True, help='Exported member database file (UTF8 Batch Export)')
ArgV.add_argument('--outcreatemembers', dest='OutputCreateMembersFile', required=True, help='CSV file for new members to (used for batch import to YM)')
ArgV.add_argument('--outupdatemembers', dest='OutputUpdateMembersFile', required=True, help='CSV file to update members (used for batch import to YM)')
ArgV.add_argument('--outdupeusernames', dest='OutputDupeUsernamesFile', required=True, help='CSV file for duplicate usernames')

#ArgV.add_argument('-f', '--format', default='csv', dest="Format", help='[list (human readable) | csv | excel (requires Output File)]')
#ArgV.add_argument('-m', '--mode', default='email', dest="Mode", help='Duplicate Mode: [name | email | member_email]')
#ArgV.add_argument('-s', '--summary', default=False, dest='Summary', help='Print count at end of output', action='store_true')

Args = ArgV.parse_args()
InputChampionsFile = Args.InputChampionsFile
InputMembersFile = Args.InputMembersFile
InputNonMembersFile = Args.InputNonMembersFile
OutputCreateMembersFile = Args.OutputCreateMembersFile
OutputUpdateMembersFile = Args.OutputUpdateMembersFile
OutputDupeUsernamesFile = Args.OutputDupeUsernamesFile


#OutFormat = Args.Format
#Mode = Args.Mode
#PrintSummary = Args.Summary

DefaultMemberTypeCode = 'Basicl'
DefaultMemberDescription = 'Basic Individual Subscription (Free)'

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

DFDepUsernames = pd.DataFrame()
DFDepNames = pd.DataFrame()
DFDefEmails = pd.DataFrame()

def check_infiles():
  if not os.path.exists(InputChampionsFile):
    print("Input file not found: " + InputChampionsFile, file=sys.stderr)
    exit(1)

  if not os.path.exists(InputChampionsFile):
    print("Input file not found: " + InputMembersFile, file=sys.stderr)
    exit(1)

  if not os.path.exists(InputChampionsFile):
    print("Input file not found: " + InputNonMembersFile, file=sys.stderr)
    exit(1)


def read_csv_file(read_filename):
  try:
    infile = open(read_filename, mode='r', encoding="utf-8")
  except IOError as e:
    print('I/O error({0}): {1}'.format(e.errno, e.strerror), file=sys.stderr)
  except: #handle other exceptions such as attribute errors
    print('Unexpected error: ', sys.exc_info()[0], file=sys.stderr)

  infile = infile.read()
  infile.close()

  return infile


def read_excel_file(read_filename):
  try:
    infile = open(read_filename, mode='r')
  except IOError as e:
    print('I/O error({0}): {1}'.format(e.errno, e.strerror), file=sys.stderr)
  except: #handle other exceptions such as attribute errors
    print('Unexpected error: ', sys.exc_info()[0], file=sys.stderr)

  infile = infile.read()
  infile.close()

  return infile


def parse_csv_data(in_csv_filename):
  dataset = pd.read_csv(in_csv_filename, dtype='unicode', encoding='unicode_escape')
  return dataset


def parse_excel_data(in_excel_filename):
  dataset = pd.read_excel(in_excel_filename)
  return dataset


def save_csv(out_csv_filename, df_export):
  df_export.to_csv(out_csv_filename, sep=',', encoding='utf-8', index=False) 


def generate_username(row):
  return row['First Name'][0:1].lower() + row['Last Name'].lower()


def generate_password():
  return secrets.token_urlsafe(8)


# Returns raw Members list that match chanpions
def find_members_by_name(in_df_champions, in_df_members):
  return in_df_members[in_df_members["Last Name"].isin(in_df_champions["Last Name"]) & in_df_members["First Name"].isin(in_df_champions["First Name"])]


# Returns raw Non-Members list that match chanpions
def find_non_members_by_name(in_df_champions, in_df_nonmembers):
  return in_df_nonmembers[in_df_nonmembers["Last Name"].isin(in_df_champions["Last Name"]) & in_df_nonmembers["First Name"].isin(in_df_champions["First Name"])]


# Returns names not listed in members or non-members
def find_non_members_by_name(in_df_champions, in_df_members, in_df_nonmembers):
  in_df_notlisted = pd.DataFrame({'First Name': '', 'Last Name': ''})  
  return in_df_notlisted( ( ~(in_df_nonmembers[in_df_nonmembers["Last Name"].isin(in_df_champions["First Name"]) ) & ~(in_df_nonmembers["First Name"].isin(in_df_champions["First Name"]) ) )



def make_new_member_df(in_non_member_df, in_member_df):
  # I'm sure there's a more elegent way to handle this, but I don't have time to learn pandas ATM

  copy_df = in_non_member_df.copy()
  copy_df['Member Type Code'] = DefaultMemberTypeCode
  copy_df['Membership'] = DefaultMemberDescription
  
  copy_df['Username'] = copy_df.apply(lambda row: generate_username(row), axis=1)
  
  dupe_username_rows = copy_df.loc[copy_df['Username'].str.contains(copy_df['Username'], case=False, na=False)]
  print(dupe_username_rows)


  copy_df['Password'] = copy_df.apply(lambda row: generate_password(), axis=1)
  
  date_now = datetime.now().strftime("%m/%d/%Y")

  copy_df['Registration Date'] = date_now
  copy_df['Date Last Renewed'] = date_now
     
  copy_df['Member Approved'] = 'Y'
     
  copy_df['Champion'] = 'Yes'

  copy_df['Member Name Title'] = copy_df[NewMemberCols['Member Name Title']]
  copy_df['Member Name Suffix'] = copy_df[NewMemberCols['Member Name Suffix']]
  copy_df['Employer Name'] = copy_df[NewMemberCols['Member Name Suffix']]
  
  ret = copy_df[NewMemberCols.keys()]
  return ret


# def find_dupes(data):
#   ret = None
  
#   #print("Mode: ", Mode)

#   if Mode == 'email':
#     ret = data[data.duplicated(['Email_Address'], keep=False)].dropna(subset=['Email_Address']).sort_values(by=['Membership', 'Email_Address'])
#   elif Mode == 'member_email':
#     # Make list of duplicate email addresses that have at least one row with a membership set
#     member_list = data[data.Membership.notnull()]
#     #member_list = data.Membership.notnull()
#     ret = data[data.isin(member_list['Email_Address']).duplicated(['Email_Address'], keep=False)].sort_values(by=['Email_Address', 'Membership']).dropna(subset=['Email_Address'])
#     #ret = data[data.isin(member_list['Email_Address'].values)].dropna(subset=['Email_Address'])
#     #ret = data.join(member_list, on='Email_Address')
#     #ret = member_list
    
#     print(tabulate(pd.DataFrame(ret, columns=['First_Name', 'Last_Name', 'Email_Address', 'Employer_Name', 'Membership', 'Web_Site_Member_ID']), showindex=False, headers=member_list.columns)).encode(sys.stdout.encoding, errors='replace').decode('ascii', 'ignore')
#     print("blah!!!")
#     exit()

#   #print(ret)
#   #exit(0)

#   return ret


# def output_dupes(data):
#   if OutFormat == 'list':
#     #pd.set_option('display.max_rows', None)

#     out_list_full = pd.DataFrame(data, columns=['First_Name', 'Last_Name', 'Email_Address', 'Employer_Name', 'Web_Site_Member_ID'])

#     # Remove duplicate entries where the email address, name, and employer only differ due to case
#     out_list_filter =  out_list_full.loc[out_list_full.apply(lambda x: x.astype(str).str.lower()).drop_duplicates(subset=['First_Name', 'Last_Name', 'Email_Address', 'Employer_Name'], keep='first').index]
    
#     print(tabulate(out_list_filter, showindex=False, headers=out_list_full.columns))
#     #print(tabulate(out_list_full, showindex=False, headers=out_list_full.columns))

#     if PrintSummary:
#       print("Duplicate records: ", len(out_list_filter.index), "/",  len(out_list_full.index))

#     return

def find_test(pd_in):
  df_found = pd_in.loc[pd_in['Email Address'].str.contains("nesta.net", case=False, na=False)]
  return df_found



def run():
  #print("Input file: ", InputFile)
  #print("Output file: ", OutputFile)

  check_infiles()
  pd_champions = parse_excel_data(InputChampionsFile)
  pd_members = parse_csv_data(InputMembersFile)
  pd_nonmembers = parse_csv_data(InputNonMembersFile)
  
  # found_test = find_test(pd_nonmembers)
  # cols_test =  make_new_member_df(found_test)
  # print(cols_test)

  # save_csv(OutputCreateMembersFile, cols_test)

  #champion_members_by_name = find_members_by_name(pd_champions, pd_members)
  #print(champion_members_by_name)
  #save_csv(OutputUpdateMembersFile, champion_members_by_name)
  print(find_non_members_by_name())


run()
