# CreateNewTeamsMeeting Macro for Outlook

This macro allows you to create a new Microsoft Teams meeting with today's start date and end time selected 7 days from today, with a single click. It also copies the body of the meeting to the clipboard.

# Requirements
1. Microsoft Outlook
2. Microsoft Teams

# Installation
Open Outlook and press ALT+F11 to open the VBA editor.
In the VBA editor, create a new module by clicking on the "Insert" menu and then selecting "Module."
Once you have a new module, copy and paste the code from the CreateNewTeamsMeeting.vb file into the module.
Save the macro by clicking on the "File" menu and then selecting "Save."
Close the VBA editor.

# Usage
In Outlook, go to the "View" tab and click on "Macros."
Select the "CreateNewTeamsMeeting" macro and click on the "Run" button.
The macro will automatically create a new Teams meeting with today's start date and end time selected 7 days from today.
The body of the meeting will be copied to the clipboard, so you can paste it in other applications.
Once the meeting is created, you can edit the title and add anything else into the body.

# Note
The macro uses the Outlook Object Model and it's a VBA code, which is only available in the Outlook desktop application and not in the web or mobile version of Outlook.
The code is only targeting the default calendar, if you want to target other calendar you need to specify their folder path.
The code is set to create the meeting with end date 7 days from today, if you want to change that you can modify the endDate = DateAdd("d", 7, startDate) this line of code to the number of days you want the meeting to end.
The code is set to copy the body of the meeting to the clipboard, this feature can be removed if you don't need it by removing the lines of codes that copy the body to the clipboard.
