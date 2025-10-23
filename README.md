# VisioParser
Extracts Visio object model from vsdx file for further processing or JSON export.  
No COM Interop, no Visio installation required.

## Usage
### Parse VSDX file and export as JSON
```C#
using Geradeaus.Visio;

var vp = new VisioParser();
vp.ParseVsdx(@"C:\Temp\Demo.vsdx");

foreach (var page in vp.VisioModel.Document.Pages)
    Console.WriteLine($"Page ID: {page.Key} Name: {page.Value.Name}");

Console.WriteLine($"{vp.VisioModel.Document.Pages.Count} Pages");

vp.ExportJson(@"C:\Temp\Demo.json");
```

### Import from JSON
```C#
using Geradeaus.Visio;

var vp = new VisioParser();
vp.ImportJson(@"C:\Temp\Demo.json");

foreach (var page in vp.VisioModel.Document.Pages)
    Console.WriteLine($"Page ID: {page.Key} Name: {page.Value.Name}");

Console.WriteLine($"{vp.VisioModel.Document.Pages.Count} Pages");
```