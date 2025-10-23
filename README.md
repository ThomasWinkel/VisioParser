# VisioParser
Extracts Visio object model from vsdx file for further processing or JSON export.

## Usage
```C#
using Geradeaus.Visio;

var vp = new VisioParser();
vp.ParseVsdx(@"C:\Temp\Demo.vsdx");

foreach (var page in vp.VisioModel.Document.Pages)
    Console.WriteLine($"Page ID: {page.Key} Name: {page.Value.Name}");

vp.ExportJson(@"C:\Temp\Demo.json");

vp.VisioModel = null;
vp.ImportJson(@"C:\Temp\Demo.json");
Console.WriteLine($"{vp.VisioModel.Document.Pages.Count} Pages");
```