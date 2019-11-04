# ExportToExcel
This library is designed as a wrapper around the fantastic [EPPlus](https://github.com/JanKallman/EPPlus). It provides a simple interface for exporting data to Excel in the form of strongly-typed models. There are three objects used in implementation, one of which is mostly optional.

## Basic Usage
The very simplest usage requires the creation of an `XlSheet` for each sheet of data you want in the Excel spreadsheet, adding those to the `IEnumerable` of your choice, and passing them to the `XlExporter`.

```csharp
// Get data from database
var empData = ctx.GetEmployees();
var volData = ctx.GetVolunteers();
var execData = ctx.GetExecutives();

// Create sheets
var empSheet = new XlSheet<Employee>("Employees", empData);
var volSheet = new XlSheet<Employee>("Volunteers", volData);
var execData = new XlSheet<Employee>("Executives", execData);

var reportData = new List<XlSheet>
{
    empSheet,
    volSheet,
    execSheet
};

var report = new XlExporter(reportData);
var excelFile = report.Run();
```

This assigns the raw file data into the variable, which can then be used to save it to disk in your preferred method, or return it to another method, say as a FileActionResult in an MVC application, triggering a download.

By providing some additional information, the exporter can save the data to a file automatically. The following saves the report as `Employee List.xlsx` in the current working directory.

```csharp
var report = new XlReporter(reportData, "Employee List.xlsx")
```

The `XlFileInfo` class can be used to encapsulate file creation information, such as a template, filename, output location, and an alternate backup location as well.
