using System;
using System.Windows.Forms;
using OfficeOpenXml;

namespace FileCollector
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            // EPPlus 5+ requires explicit license context. Set it before any worksheet is touched.
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
