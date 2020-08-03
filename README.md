# Mark attendance of Google meet attendees
-an apps script that marks attendance in a Google spreadsheet in the format generally used in schools

### Attendance is marked as Present and Absent against nominal roll after importing from Chrome extension

## How to use ? Follow exactly as below

## Step 1 Install Chrome extension 
(To be done before starting the G-meet session)

1) Install Chrome extension Meet Attendance from https://chrome.google.com/webstore/detail/meet-attendance/nenibigflkdikhamlnekfppbganmojlg?hl=en

Click on Add to Chrome > Add extension. 

2) For the first time you use it in your class, take attendance by clicking on the show everyone button and then click People.

3) Click on the icon to open the spreadsheet. Rename the spreadsheet if required. 

3) You can add sheets and rename them (at the bottom). Have a sheet for each of your papers.

Points to note: 

All future attendees will be captured in this spreadsheet. 

Attendance will be taken on the sheet which is in the first position -so drag the sheet you want attendance in to first position as per your session.

Add first two columns: The first column is Names of students (as per their email), second column is Roll nos.

From the third column you'll see the names of attendees who came for each meet session at the time you took attendance.

Ensure that the first two columns have ALL the expected attendees (sorted by Roll number-if that is how you need to submit it). Their names must be EXACTLY as per their email account name (and not official records)

Each time you click on the Show everyone button > then People during a meet session, it will fill a new column with attendees at that time. You can turn this off by toggling off the horizontal toggle button.

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

Run the Mark Attendance program each time after you get the attendees using the extension 

OR you can run the program after many sessions.

Attendance will be taken on the sheet which is in the first position -so drag the sheet you want attendance in to first position as per your session.

