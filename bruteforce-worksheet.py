import win32com.client
import time

excel_file = r'C:\Code\YouTube\python-brute-force-excel-password\Book1.xlsx'
password_file = r'C:\Code\YouTube\python-brute-force-excel-password\passwords.txt'

excel = win32com.client.Dispatch('Excel.Application')

password_list = []

# extract passwords from file and load to list object
with open(password_file, 'r', encoding='utf-8') as pwd:
    passwords = pwd.readlines()
    for password in passwords:
        password_list.append(password.replace('\n', ''))

wb = excel.Workbooks.Open(excel_file, False, True, None)

for password in password_list:
    try:
        wb = excel.Workbooks.Open(excel_file, False, True, None)
        ws = wb.Sheets(1)
        ws.Unprotect(password)
        print('Successfully Password: ', password)
        excel.Quit()
        time.sleep(1)
        quit()
    except:
        continue

print('Password not able to unlock Worksheet')