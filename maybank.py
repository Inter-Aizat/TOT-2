
import os, sys

from win32com.client import Dispatch

excel = Dispatch("Excel.Application")
excel.Visible = 1
excel.DisplayAlerts = False

xlUp = -4162

report = excel.Workbooks.Open(r'C:\Users\aizat\Desktop\TOT 2\INPUT\MAYBANK\email_4B7CC1A7F2BFD1F7968744ABB3914C18.xlsx', UpdateLinks = 0)
reportData = report.Worksheets(1)
reportRow = reportData.Cells(reportData.Rows.Count, 1).End(xlUp).Row

for i in range(1,reportRow+1):
    status_data = reportData.Cells(i,11).Value
    if status_data == "Delivered":
        continue
    print(reportData.Cells(i,15).Value)
    # print(reportData.Cells(i,11).Value)
# print(reportData.Cells(1,2).Value)

# reportData.Cells(1,2).Value

# excel.quit()

# if getattr(sys, 'frozen', False):
#     application_path = os.path.dirname(sys.executable)
# elif __file__:
#     application_path = os.path.dirname(__file__)

# main_dir = application_path.split('\\')[-1]

# for dirs, sub_dir, files in os.walk(application_path):
#     if 'venv' in sub_dir or '.git' in sub_dir or 'build' in sub_dir or 'dist' in sub_dir:
#         sub_dir.remove('venv')
#         sub_dir.remove('.git')
#         sub_dir.remove('build')
#         sub_dir.remove('dist')
#     if dirs.split('\\')[-1] == main_dir:
#         continue
#     if "MAYBANK" in dirs:
#         for filenames in files:
#             file_path = os.path.join(dirs,filenames)
#             # print(file_path)
#             with open(file_path) as file:
#                 print(len(file.readlines()))
    
# os.system('pause')