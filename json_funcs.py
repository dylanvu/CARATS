# json_funcs.py

"""
all functions relevant to json and json objects
"""


def json_print(obj):
    """
    takes in a json object and prints out a "neat" text version of the json
    useful for viewing and/or debugging json objects
    :param obj:
    :return:
    """

    import json  # import json library to utilize dumps function

    text = json.dumps(obj, sort_keys=True, indent=4)  # create a text string of the json object
    print(text)  # display formatted text


def get_req_api_json(quarter, deptname, classnum):
    """
    performs API get request
    :param quarter: current quarter
    :param deptname: name of course department, part of course ID
    :param classnum: class code, part of course ID
    :return: resp.json(): json object from api call
    """

    import requests     # import requests library to make API calls
    from dotenv import dotenv_values

    # constants for all get requests
    base_URL = "https://api.ucsb.edu/academics/curriculums/v3/classes/search"    # base URL

    headers = dotenv_values(".env")

    # check if ALL courses from department were requested
    if classnum == "ALL":
        parameters = {
            "quarter": quarter,         # current quarter
            "subjectCode": deptname,    # department name
            "objLevelCode": "U",        # undergraduate courses
            "pageSize": 500             # max query size
        }
        resp = requests.get(base_URL, params=parameters, headers=headers)   # API get request
        return resp.json()  # return json object of response from get request
    else:
        parameters = {
            "quarter": quarter,                     # current quarter
            "courseId": deptname + " " + classnum   # course ID from department
        }
        resp = requests.get(base_URL, params=parameters, headers=headers)   # API get request
        return resp.json()  # return json object of response from get request


def json_data_parse(json_obj):
    """
    parse through json data of a given json object
    :param json_obj: json object of classes and the relevant information
    :return: data_list: list of all relevant extracted data from json_object
    """

    data_list = []  # initialize data_list to append extracted data to
    for crsIdx in range(0, json_obj["total"]):            # loop through all courses
        course = json_obj["classes"][crsIdx]              # search for course
        crsString = " ".join(course["courseId"].split())  # course ID

        for lec in course["classSections"]:               # loops through all lectures (and sectionss)
            enrlCd = lec["enrollCode"]                    # enrollment code

            # professors
            profString = ""         # initialize professor string
            checkLecSecString = ""  # initialize lecture vs. section checking string
            for prof in lec["instructors"]:  # loop through all instructors
                profString += (prof["instructor"] + ", ")   # add professor to string and break between professors
                checkLecSecString += prof["functionCode"]   # check the professor's functionality in the lecture
            profString = profString[:-2]    # remove the last break between professors
            if profString == "":    # check for blank professor strings
                profString = "TBA"  # replace blank strings with a placeholder to avoid empty strings

            # day and time
            timeString = ""  # initialize time string
            for timeLoc in lec["timeLocations"]:    # loop through all time locations
                tl = timeLoc["days"]        # collect lecture days
                bt = timeLoc["beginTime"]   # collect lecture start times
                if tl is None:              # check for NoneType objects in lecture days
                    tl = "TBA"              # replace NoneType with placeholder to avoid empty strings
                else:                       # otherwise proceed as normal
                    tl.replace(" ", "")     # remove spaces in the string
                if bt is None:              # check for NoneType objects in lecture start times
                    bt = "TBA"              # replace NoneType with placeholder to avoid empty strings
                else:                       # otherwise proceed as normal
                    bt.replace(" ", "")     # remove spaces in the string
                timeString = tl + " " + bt  # concatenate both time strings

            # temporal data: numbers
            enrlTot = lec["enrolledTotal"]          # extract total enrollment
            maxSpace = lec["maxEnroll"]             # extract maximum space
            enrlPct = 0                             # initialize percent available value
            if enrlTot is None:                     # check for NoneType objects in total enrollment
                enrlTot = 0                         # replace NoneType with 0
            if maxSpace is None:                    # check for NoneType objects in maximum enrollment
                maxSpace = 0                        # replace NoneType with 0
            else:                                   # otherwise, proceed as normal
                enrlPct = 1 - (enrlTot / maxSpace)  # calculate percent space left
            enrlLeft = maxSpace - enrlTot           # calculate number of spaces left
            if enrlTot > maxSpace:                  # what to do if total enrolled exceeds maximum space
                enrlLeft = 0                  # set total enrolled equal to maximum space

            if "Teaching and in charge" in checkLecSecString:   # append lectures
                data_list.append([f'{enrlCd}', crsString, profString, timeString, enrlLeft, enrlPct, maxSpace])

    return data_list
