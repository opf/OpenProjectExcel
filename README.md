# OpenProjectExcel

# HowTo

How To

There are 6 sheets in the workbook. „Workpackages“ is the main sheet. The other sheets are just cache and configuration and are hidden by default.

By opening the workbook (after accepting macros etc.) a window will be shown. You have to define a URL (including the „http“-stuff), an API key/ token (found in „my account“ in your OpenProject instance) and a project (use semantic id, not the numeric id). So far the input is mandatory, the QueryID is not.

If defined, you are able to load default values, which are stored in the worksheet „Default“. Please note, that your API token is stored in this worksheet as well. Check this before sending the Excel-file to anyone. Press „Accept“ to initialize.

After pressing „Accept” you will get an empty worksheet “Workpackages”. You can define the attributes you want to download or upload by choosing them from the dropdown menu in row one. It is possible to define as many attributes/ columns as you want.

To perform a download or an upload press Ctrl + B. By pressing, a userform will be shown to download workpackages as defined (project/ QueryID). Same procedure for uploading the changed data. You can change your chosen project (and URL etc.) by pressing „Show chosen project“.
