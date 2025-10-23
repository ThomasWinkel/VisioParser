using Geradeaus.Visio;
using System.Diagnostics;

// Serial performance test
Stopwatch stopwatch = Stopwatch.StartNew();

var vp1 = new VisioParser();
vp1.ParseVsdx(@"C:\Temp\vsdx\vsdx1.vsdx");
vp1.ExportJson(@"C:\Temp\vsdx\vsdx1.json");

var vp2 = new VisioParser();
vp2.ParseVsdx(@"C:\Temp\vsdx\vsdx2.vsdx");
vp2.ExportJson(@"C:\Temp\vsdx\vsdx2.json");

stopwatch.Stop();
Console.WriteLine($"Seriell: {stopwatch.Elapsed.TotalSeconds:F3} Sekunden");

// Parallel performance test
stopwatch = Stopwatch.StartNew();

await Task.WhenAll(
    Task.Run(() =>
    {
        var vp3 = new VisioParser();
        vp3.ParseVsdx(@"C:\Temp\vsdx\vsdx1.vsdx");
        vp3.ExportJson(@"C:\Temp\vsdx\vsdx1.json");
    }),
    Task.Run(() =>
    {
        var vp4 = new VisioParser();
        vp4.ParseVsdx(@"C:\Temp\vsdx\vsdx2.vsdx");
        vp4.ExportJson(@"C:\Temp\vsdx\vsdx2.json");
    })
);

stopwatch.Stop();
Console.WriteLine($"Parallel: {stopwatch.Elapsed.TotalSeconds:F3} Sekunden");

// Object model demo
foreach (var page in vp1.VisioModel.Document.Pages)
    Console.WriteLine($"Page ID: {page.Key} Name: {page.Value.Name}");

// Import JSON demo
vp1.VisioModel = null;
vp1.ImportJson(@"C:\Temp\vsdx\vsdx1.json");
Console.WriteLine($"{vp1.VisioModel.Document.Pages.Count} Pages");