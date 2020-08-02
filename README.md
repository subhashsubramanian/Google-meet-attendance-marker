# Mark attendance of Google meet attendees
-an apps script that marks attendance in a Google spreadsheet in the format generally used in schools

### Attendance is marked as Present and Absent against nominal roll after importing from Chrome extension

## How to use ?

## Step 1 Chrome extension 

Install Chrome add on Meet Attendance from https://chrome.google.com/webstore/detail/meet-attendance/nenibigflkdikhamlnekfppbganmojlg?hl=en

For the first time you use it in your class, take attendance by clicking on the button and then click on the icon to open the spreadsheet

You can choose to pull all future attendees in this same spreadsheet- you can even rename it.

The first column is Names of Attendees, second column is emails, from third column you'll see the names of attendees who came for each meeting. Each column will give you the attendees in that meeting when you took attendance.

Ensure that the first column has ALL the expected attendees in the roll (sorted by Roll number-if that is how you need to submit it). If not you can add their names EXACTLY as per their email account name

Each time you go to the meet attendance button (under People) during a meet session, it will fill a column with attendees at that time.

## Step 2 Install the Google apps script code

Open Tools  Script Editor and replace any existing code with the code in meetAttendance.gs

Rename Code.gs to meetAttendance.gs

Give the Untitled project (on top left) a name- like MarkAttendance

Refresh the spreadsheet and click on the spreadsheet menu item "Attendance" > Mark Attendance

You will be required to authorise the code and give it permissions. Do so.

Refresh the spreadsheet. You will see a menu item "Attendance". 

## Step 3 Running the Mark Attendance program

Click on Attendance > Mark Attendance

You can choose to run the Mark Attendance program each time after you get the attendees using the extension or you can run it after many meet session attendees are imported into this spreadsheet.
