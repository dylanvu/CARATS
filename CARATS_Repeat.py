# CARATS_Repeat.py

"""
runs code to be repeated throughout all three passtimes
"""

# imports
from datetime import datetime
from excel_funcs import extract_courses, time_excel_storage
from json_funcs import get_req_api_json

base_dir = r"C:\Abhiram\Abhiram UCSB\Orgs - AIChE\Web Scraper\CARATS_v4\\"

csv_fname = base_dir + "ClassList.csv"                                  # name of csv file containing departments and their courses
quarter = 20212                                              # current quarter
data_fext = ".xlsx"                                          # file extension of CARATS spreadsheet
data_fname = base_dir + "Prelim_Data_Qtr_" + str(quarter) + data_fext   # title of file to write to

course_dept_list = extract_courses(csv_fname)   # list of tuples of all departments and their courses

# create list of json objects
print("Requesting API data.")
json_list = []  # initialize json_list
for dept, courseNum in course_dept_list:    # loop through each department/course number tuple
    json_list.append(get_req_api_json(quarter, dept, courseNum))    # make api call, store to json_list

# get current time and make it a header for the sheet
print("Adding current time stamp.")
curr_time_str = datetime.now().strftime("%m/%d/%Y %H:%M")

time_excel_storage(data_fname, curr_time_str, json_list)

print("Done! :)")

