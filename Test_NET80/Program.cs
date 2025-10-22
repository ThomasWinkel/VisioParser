using Geradeaus.Visio;
using System.Diagnostics;

Stopwatch stopwatch = Stopwatch.StartNew();

var vp1 = new VisioParser(@"C:\Temp\vsdx\vsdx1.vsdx");
vp1.Parse();
vp1.ExportJson(@"C:\Temp\vsdx\vsdx1a.json");
foreach (var page in vp1.VisioModel.Document.Pages)
    Console.WriteLine($"Page ID: {page.Key} Name: {page.Value.Name}");

var vp2 = new VisioParser(@"C:\Temp\vsdx\vsdx2.vsdx");
vp2.Parse();
vp2.ExportJson(@"C:\Temp\vsdx\vsdx2a.json");

stopwatch.Stop();
Console.WriteLine($"Benötigte Zeit: {stopwatch.Elapsed.TotalSeconds:F3} Sekunden");



stopwatch = Stopwatch.StartNew();

await Task.WhenAll(
    Task.Run(() =>
    {
        var vp3 = new VisioParser(@"C:\Temp\vsdx\vsdx1.vsdx");
        vp3.Parse();
        vp3.ExportJson(@"C:\Temp\vsdx\vsdx1b.json");
    }),
    Task.Run(() =>
    {
        var vp4 = new VisioParser(@"C:\Temp\vsdx\vsdx2.vsdx");
        vp4.Parse();
        vp4.ExportJson(@"C:\Temp\vsdx\vsdx2b.json");
    })
);

stopwatch.Stop();
Console.WriteLine($"Benötigte Zeit: {stopwatch.Elapsed.TotalSeconds:F3} Sekunden");