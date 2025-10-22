# VisioParser
Extracts Visio object model from vsdx file for further processing or JSON export.

## Usage
```C#
using Geradeaus.Visio;

var vp = new VisioParser(@"C:\Temp\Demo.vsdx");
vp.Parse();

foreach (var page in vp1.VisioModel.Document.Pages)
    Console.WriteLine($"Page ID: {page.Key} Name: {page.Value.Name}");

vp.ExportJson(@"C:\Temp\Demo.json");
```