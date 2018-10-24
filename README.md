# OpenProjectExcel

# HowTo

There are 6 sheets in the workbook.

The sheet „Workpackages“ is most important. The others are just cache and configuration. Thats why theese are hidden by default.

By opening the workbook (after accepting macros ect.) a userform will be shown. You have to define a URL (including the „http“-stuff, a API key/ token (found in „my account“ in your OpenProject instance), a project (please use semantic id, not the number). Theese inputs ara mandatory. Additionaly you can define a QueryID.

You can load default values, if defined. The data is stored in worksheet „Default“. Please note, that your API token is stored there as well. So check this before sending the Excel-file someone else.
Press „accept“ to initialize.

Now you’ll get the worksheet „Workpackages“, normally empty. You can define the attributes you want to download or upload by chossing them from the dropdown in row one. You can define as many attributes/ columns as you want.

To download or upload press ctrl + b. A userform will be shown, were you can download workpackages as defined (project/ QueryID), same procedure für uploading changed data. You can change your choosen project (and URL ect.) by pressing „Show choosen project“.
