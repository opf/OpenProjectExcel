# OpenProjectExcel

Excel workbook for syncing tasks with OpenProject. Download work packages into a sheet, edit them, and upload the changes back. Filters, hierarchy, and relations are included.

This is a community project and is not officially supported by OpenProject. It is a fork of [opf/OpenProjectExcel](https://github.com/opf/OpenProjectExcel), licensed under the MIT License. Copyright (c) 2018 bpetraus.

Open `OpenProjectAPI.xlsm` in Excel for Windows and enable macros. Press `Ctrl+B` to open synchronization (download and upload). The workbook talks to the OpenProject [API v3](https://www.openproject.org/docs/api/) with an API token from your OpenProject account.

## Compatibility

| Name | Version |
| --- | --- |
| OpenProject | 14, 15, 16, **17.4** (current compatible) |
| OpenAPI | v3 |
| Operating system | Windows 10 |
| Excel | 16.0 |
| VBA | 7.1 |

Excel libraries used by the workbook:

- Visual Basic For Applications
- Microsoft Excel 16.0 Object Library
- Microsoft Forms 2.0 Object Library
- Microsoft Scripting Runtime
- Microsoft WinHTTP Services, version 5.1

## Shortcuts

- `Ctrl+B` — OpenProject synchronization (download, upload, and related actions)
- `Ctrl+N` — extra filter and sort options for the active worksheet

## How to

Blog article (English): [Synchronize Excel with OpenProject](https://www.openproject.org/synchronize-excel-openproject/)

Tutorial videos (English):

- [Basics: download, update, upload](https://www.youtube.com/watch?v=ZPK-l-e-lKI)
- [Custom queries](https://www.youtube.com/watch?v=86X334cg808)
- [Hierarchy](https://www.youtube.com/watch?v=17yqO0uXE5k)
- [Relations](https://www.youtube.com/watch?v=y7QBspWlQho)

Tutorial videos (German):

- [Basics: download, update, upload](https://www.youtube.com/watch?v=kL4u7ew0yHk)
- [Custom queries](https://www.youtube.com/watch?v=Fs-luE7gSXA)
- [Hierarchy](https://www.youtube.com/watch?v=SaCD8B9-4tM)
- [Relations](https://www.youtube.com/watch?v=wjTdXBMsrwk)

## Debugging

Turn on debugging in the user form:

![Debugging checkbox in the user form](https://user-images.githubusercontent.com/38315559/138328308-b064946b-1050-4326-9199-9191d82e7cf6.png)

The workbook then writes `OpenProjectExcelLog.txt` to your Downloads folder (`C:\Users\<user name>\Downloads`).
