# Mark attendance of Google meet attendees
-an apps script that marks attendance in a Google spreadsheet in the format generally used in schools

### Attendance is marked as Present and Absent against nominal roll after importing from Chrome extension

## How to use ? Follow exactly as below

## Step 1 Install Chrome extension 
(To be done before starting the G-meet session)

1) Install Chrome extension Meet Attendance from https://chrome.google.com/webstore/detail/meet-attendance/nenibigflkdikhamlnekfppbganmojlg?hl=en

Click on Add to Chrome > Add extension. 

2) For the first time you use it in your class, take attendance by clicking on the show everyone button and then click People.

3) Click on the tick mark icon to open the spreadsheet. Rename the file if required. 

After taking attendance the first time, you can change the first part of the sheet name (mm/dd/yyyy) to, for e.g., your subject code. But you must leave the second part - i.e. the last 10 alphabets of the sheet name - as it is. This is where all future meet attendance will be aggregated. For e.g. your sheet might be named something like 101 abcdefghij (where 101 is paper code and abcdefghij is the G-meet link for the subject). 

Edit the first column: The first column is Names of students (as per their email), add a column B with Roll no. you can delete the other columns (C, D etc.) You can give a header to column A as student Name and to column B as Roll no. 

For each subsequent meet session the names of attendees who came for that meet session will be in a new sheet named as mm/dd/yyyy abcdefghij (where abcdefghij is the meet link). You can take attendance multiple times during a meet session. It will add attendees who joined later also in the same sheet in the later rows.

Ensure that the first column has ALL the expected attendees (sorted by Roll number-if that is how you need to submit it). Their names must be EXACTLY as per their email account name (and not official records). You can copy paste names from attendees in say a meet session where all expected attendees were present.

Each time you click on the Show everyone button > then People during a meet session, it will fill a new sheet with attendees at that time. 

## Step 2 Install the Google apps script code in the spreadsheet

You must do this step after only a meet session has ended. 

1) Go to https://github.com/subhashsubramanian/Google-meet-attendance-marker/blob/master/meetAttendance.gs Click on meetattendance.gs 

2) Click on Raw, and Ctrl A (select all) then Ctrl C (Copy)

3) Open the spreadsheet created by the extension, click Tools > Script Editor > delete any existing code there and paste (Ctrl V) the copied code from meetAttendance.gs

4) Click save. Rename Code.gs to meetAttendance.gs

5) Give the Untitled project (on top left) a name- e.g. MarkAttendance

6) Refresh the spreadsheet. You will see a menu item "Attendance". 

7) click on the spreadsheet menu item "Attendance" > Mark Attendance

7) You will be required to authorise the code and give it permissions. Click l Authorise/ Allow. 

Ensure that you see menu item "Attendance"

## Step 3 Running the Mark Attendance program from the spreadsheet

Click on Attendance > Mark Attendance

Points to note:

Run the Mark Attendance program each time after you get the attendees using the extension OR you can run the program after many sessions.

### Program must be run on the sheet which on which you want attendance to be aggregated for e.g. in the sheet named 101 abcdefghij (where abcdefghij is the meet link for 101). That sheet will then have the attendance of all meet sessions in Present Absent format.

