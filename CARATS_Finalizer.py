from cleaning_funcs import *


#base_dir = r"C:\Abhiram\Abhiram UCSB\Orgs - AIChE\Web Scraper\CARATS_v4\\"

quarter = 20212                                              # current quarter
data_fext = ".xlsx"                                          # file extension of CARATS spreadsheet
final_fext = ".xlsm"
data_fname = "Prelim_Data_Qtr_" + str(quarter) + data_fext   # title of file to write to
final_fname = "CARATS_Final_Qtr_" + str(quarter) + final_fext

# condense lab courses to declutter final spreadsheet
print("Condensing courses.")
condense_labs(data_fname, final_fname)

# inject code into vba
VBA_Macro_Code = "VBA_Macro_Code.txt"
VBA_Button_Code = "VBA_Button_Code.txt"
inject_macro(final_fname, VBA_Macro_Code, VBA_Button_Code)

print("Done! :)")

"""# Valuable resource:
# Authentication: https://pythonhosted.org/PyDrive/quickstart.html

from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from pathlib import Path

path = "Condensing Result.xlsm"

path = str(Path.cwd()) + r"/" + path

gauth = GoogleAuth()
gauth.LocalWebserverAuth()  # Creates local webserver and auto handles authentication.

drive = GoogleDrive(gauth)
file = drive.CreateFile({"title": "Condensing Result.xlsm"})

file.SetContentFile(path)
file.Upload()
file = None"""
