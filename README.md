# Mark attendance of Google meet attendees
-an apps script that marks attendance in a Google spreadsheet in the format generally used in schools

### Attendance is marked as Present and Absent against nominal roll after importing from Chrome extension

## How to use ?

## Step 1 Chrome extension 

Install Chrome add on Meet Attendance from https://chrome.google.com/webstore/detail/meet-attendance/nenibigflkdikhamlnekfppbganmojlg?hl=en

For the first time you use it in your class, take attendance by clicking on the show everyone button and then click People.

Click on the icon to open the spreadsheet. Rename the spreadsheet if required. 

You can rename the sheets (at the bottom) and have a sheet for each of your papers.

All future attendees will be captured in this spreadsheet. 

Attendance will be taken on the sheet which is in the first position -drag to first position as per your session.

Note: Add first two columns: The first column is Names of Attendees (as per email), second column is Roll nos.

From the third column you'll see the names of attendees who came for each meeting. Each column will give you the attendees in that meeting when you took attendance.

Ensure that the first two columns have ALL the expected attendees in the roll (sorted by Roll number-if that is how you need to submit it). Their names must be EXACTLY as per their email account name (and not official records)

Each time you go to the Show everyone button and People during a meet session, it will fill a new column with attendees at that time.

## Step 2 Install the Google apps script code in the spreadsheet

You must do this step after only a meet session has ended. Click on meetattendance.gs 

Click on Raw, and Ctrl A (select all) then Ctrl C (Copy)

Open the spreadsheet created by the extension, click Tools > Script Editor > replace any existing code there by pasting (Ctrl V) the copied code from meetAttendance.gs

Click save. Rename Code.gs to meetAttendance.gs

Give the Untitled project (on top left) a name- e.g. MarkAttendance

Refresh the spreadsheet and click on the spreadsheet menu item "Attendance" > Mark Attendance

You will be required to authorise the code and give it permissions. Authorise/ Allow. Do so.

Refresh the spreadsheet. You will see a menu item "Attendance". 

## Step 3 Running the Mark Attendance program from the spreadsheet

Click on Attendance > Mark Attendance

Run the Mark Attendance program each time after you get the attendees using the extension 

OR you can run the program after many sessions.
