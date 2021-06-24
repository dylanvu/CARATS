# CARATS

Welcome to CARATS! This excel sheet logs the availability of select UCSB classes so you can see what courses, times, or professors fill up first. Over time, you can see what courses/times/professors filled up first. Hopefully, this helps you (the user) gauge what courses to prioritize first during your first and second passtimes and help you plan ahead easier.

First, please make sure you have enabled macros. At the top, there should be a bar at the top of this worksheet that asks to enable macros. If macros are still not enabled, follow the instructions to "enable macros in all workbooks via Trust Center" here:
https://www.ablebits.com/office-addins-blog/2020/03/11/enable-disable-macros-excel/#:~:text=To%20get%20macros%20enabled%20in,all%20macros%20and%20click%20OK.

Click on the "View Course" button to see graphs of courses that we have data for. Type in the abbreviated course department name and the course number. For example, type "CH E 110A". Do NOT type "Chemical Engineering 110A", "Chemical Engineering Thermodynamics", "CHE110A", etc as these will all generate errors. There will be three graphs (Space Left, Percentage, Total Space). Drag the charts out so that you can see all three plots.

Excel may add extra characters in courses that have spaces in them. For example, "CH E" may turn into something like "CH+$7:$20E." This happens because you are holding down shift and pressing down space while typing in the course name. Do not hold down shift when pressing the spacebar. Just delete this extra portion and continue where you left off.

Thank you and we hope you find this tool useful!

# Developers
* Setting up the API Key
    1. Obtain the API Key
    2. Create a .env file
    3. Inside the .env file, enter the following: `ucsb-api-key="API_KEY_HERE"` where `API_KEY_HERE` refers to the API key. Note that it is inside quotations.
    4. Make sure python-dotenv is installed via pip