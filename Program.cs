using System;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace ExcelComApps
{
    public class Program
    {
        [STAThreadAttribute]
        static void Main(string[] args)
        {
            var obj = new ExcelProcessApps();
            var p = string.Empty;
            var l = string.Empty; // Path.GetDirectoryName(p);
            if (args == null)
            {

            }
            if (args.Length == 1)
            {
                p = args[0];
            }
            if (args.Length == 2)
            {
                p = args[1];
            }
            //  Console.WriteLine($"start time: {DateTime.Now}");
            obj.chartImageExtension = ".SVG";
#if DEBUG
            p = @"C:\Users\leaflet_vba_delhi\Desktop\AmazonShare\temp\Legend\12345.xlsx";
          //  obj.chartImageExtension = ".png";
#endif
            l = Path.GetDirectoryName(p);
            try
            {
                obj.GetAllTableImageChartInformation(p, l);
            }
            catch
            {

            }
            finally
            {
#if DEBUG
#else
                obj.Dispose();
#endif

            }
            try
            {

#if DEBUG
#else
                var proId = GlobalsPoint.AfterAppPID.Except(GlobalsPoint.BeforeAppPID).FirstOrDefault();
                Process processes = Process.GetProcessById(proId);
                processes.Kill();
#endif
            }
            catch { }

            //  Console.WriteLine($"end time: {DateTime.Now}");
            //  Console.ReadLine();
        }
    }
}
