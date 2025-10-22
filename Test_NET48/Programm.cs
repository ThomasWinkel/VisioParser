using Geradeaus.Visio;
using System;
using System.Diagnostics;
using System.Threading.Tasks;

namespace Test_NET48
{
    internal class Programm
    {
        static async Task Main(string[] args)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();

            var vp1 = new VisioParser(@"C:\Temp\vsdx\vsdx1.vsdx");
            vp1.Parse();
            vp1.ExportJson(@"C:\Temp\vsdx\vsdx1a.json");

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
        }
    }
}
