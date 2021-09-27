import os
import xlsxwriter
import sys, json
from datetime import date, datetime

def search(list, search_term):
    for i in range(len(list)):
        if list[i] == search_term:
            return True
    return False
jobid = sys.argv[1]
employee = json.loads(sys.argv[2])
jobdate = json.loads(sys.argv[3])
reports = json.loads(sys.argv[4])
file_name = sys.argv[5]
cell_names = ['C1', 'D1', 'E1', 'F1']
eval_names = ["Travel In","Team Arrival","Time In","Time Out"]
location = os.path.join(os.path.expanduser('~'), 'Documents', 'Job_Timesheets')
isExists = os.path.exists(location)
if not isExists:
    os.makedirs(location)
outWorkbook = xlsxwriter.Workbook(os.path.join(location, file_name))
row = 1
bold = outWorkbook.add_format({'bold': True})
time_format = outWorkbook.add_format({'num_format': 'hh:mm AM/PM'})
duration_format = outWorkbook.add_format({'num_format': 'hh:mm:ss'})
date_format = outWorkbook.add_format({'num_format': 'mm/dd/yyyy'})
sheet_names = []

def search(list, search_term):
    for i in range(len(list)):
        if list[i] == search_term:
            return True
    return False

for dateIndex in jobdate:
    sheet_names.append("{}".format(dateIndex))

for name in sheet_names:
    print('sheet name:', name)
    outSheet = outWorkbook.add_worksheet("{}".format(name))
    getSheet = outWorkbook.get_worksheet_by_name("{}".format(outSheet.get_name()))
    if search(sheet_names, getSheet.get_name()):
        print("worksheet was found: ", getSheet.get_name())
        getSheet.write("A1", "Name", bold)
        getSheet.write("B1", "Date", bold)
        getSheet.write("G1", "Travel Time", bold)
        getSheet.write("H1", "Project Hours", bold)
        getSheet.write("I1", "Total Time", bold)
        day_reports = reports[getSheet.get_name()]
        for rep in day_reports:
            emp_dict = dict(rep['teamMember'])
            eval_list = list(rep['evaluationLogs'])
            filtered_evals = list(filter(lambda c: c['label'] != 'Total Time', eval_list))
            print(filtered_evals)
            eval_dict = {
                "Travel In": datetime.strptime(filtered_evals[0]['value'], '%m-%d-%Y %H:%M:%S'),
                "Team Arrival": datetime.strptime(filtered_evals[1]['value'], '%m-%d-%Y %H:%M:%S'),
                "Time In": datetime.strptime(filtered_evals[1]['value'], '%m-%d-%Y %H:%M:%S'),
                "Time Out": datetime.strptime(filtered_evals[2]['value'], '%m-%d-%Y %H:%M:%S')
            }
            getSheet.write(row, 0, emp_dict['name'])
            getSheet.set_column_pixels(0, 0, 120)
            date_time = datetime.strptime(rep['date'], '%m-%d-%Y')
            getSheet.write_datetime(row, 1, date_time, date_format)
            getSheet.set_column_pixels('B:I', 80)
            for index, item in enumerate(eval_names):
                getSheet.write(cell_names[index], item, bold)

            for key, eval_log in eval_dict.items():
                index = list(eval_dict).index(key)
                getSheet.write_datetime(row, index + 2, eval_log, time_format)

            getSheet.write_formula(row, 6, '=TEXT(D%d-C%d, "hh:mm:ss")' % (row + 1, row + 1))
            getSheet.write_formula(row, 7, '=TEXT(F%d-E%d, "hh:mm:ss")' % (row + 1, row + 1))
            getSheet.write_formula(row, 8, '=SUM(H%d+G%d)' % (row + 1, row + 1), duration_format)
            if len(day_reports) > 1:
                row += 1
outWorkbook.close()