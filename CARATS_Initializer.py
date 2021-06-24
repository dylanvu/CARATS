# CARATS_Initializer.py
"""
initializes spreadsheet to contain all relevant info for the quarter's passtimes
"""

# imports
from excel_funcs import extract_courses, init_excel_storage
from json_funcs import get_req_api_json
from os import path

csv_fname = "ClassList.csv"                                  # name of csv file containing departments and their courses
quarter = 20212                                              # current quarter
data_fext = ".xlsx"                                          # file extension of CARATS spreadsheet
data_fname = "Prelim_Data_Qtr_" + str(quarter) + data_fext   # title of file to write to

# checking whether appropriate file already exists
print("Checking File Name.")
if path.isfile(data_fname):
    print("A file with this name already exists. Please change the file name and rerun this program.")
else:

    print("Extracting list of courses.")
    course_dept_list = extract_courses(csv_fname)   # list of tuples of all departments and their courses

    # create list of json objects
    print("Requesting API data.")
    json_list = []  # initialize json_list
    for dept, courseNum in course_dept_list:    # loop through each department/course number tuple
        json_list.append(get_req_api_json(quarter, dept, courseNum))    # make api call, store to json_list

    print("Initializing and populating Excel sheet.")
    init_excel_storage(data_fname, json_list)  # create excel sheet

    print("Done! :)")
