#An-Automated-Email-Reminder-System

**Overview**

This is an automated email reminder system developed using a simple algorithm and Google Workspace tools such as Spreadsheet and AppScript. It can be used to manage Non-Closed (NC) items, project submissions, document submissions, and other applications that require periodic reminders. The tool simplifies the process of sending reminder emails and saves time and effort for the users.

**How to Use**

To use the tool, please make a copy of the Spreadsheet provided here :
https://docs.google.com/spreadsheets/d/1ZXC8O1AbQhh0zJP8Z4KPwv2krd6qhnc_tE-0TeBhJjE/edit?usp=sharing

Sheet 1 is the main sheet that includes all project or NC details, which can be filled in manually. It includes essential information such as name, email, project name, project description, open date, due date, and a status tick box (complete or not complete). The tool will only send emails for unticked status projects. If a project is complete, the status should be changed to a tick status manually. Sheet 2 provides clear information about each person and their project, highlighting overdue projects with an asterisk symbol.

**The tool has several features, including:**

- It sends two reminder emails (two days before the due date and on the due date).
- It works entirely on automation and can also be used manually by clicking a single button.
- Sheet 2 is automatically updated every time, containing up-to-date information on each person, their project, and any outdated projects.

The tool can work in two ways: manual and automated. The automated version triggers an email every day at 8 am and updates the list in Sheet 2. A custom menu has been added to the top of the tool to allow for manual use.

**Background**

During my internship at Valeo, I noticed that reminder emails for Non-Closed (NC) items were being sent manually, which was time-consuming and prone to errors. To address this issue, I planned to develop an efficient and user-friendly email reminder tool using a simple algorithm to make the process faster, along with Google Workspace tools such as Spreadsheet and AppScript. Although my colleague initiated the process of automating this task at Valeo, I developed this tool independently and implemented it in my way for my purpose.

**Conclusion**

Overall, this tool is an effective way to manage Non-Closed NCs, project submissions, document submissions, and other applications. It simplifies the process of sending reminder emails and saves time and effort for the users.
