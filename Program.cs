using System;
using System.Windows.Forms;

namespace FileCollector
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            // Анализируем цвета логотипа
            ColorAnalyzer.AnalyzeLogo();
            
            Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}