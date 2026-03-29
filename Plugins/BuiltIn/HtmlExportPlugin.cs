using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace FileCollector.Plugins.BuiltIn
{
    public class HtmlExportPlugin : IExportPlugin
    {
        private IPluginHost? host;

        public string Name => "HTML Exporter";
        public string Version => "1.0";
        public string Description => "Экспорт результатов в HTML отчет";
        public string FormatName => "HTML";

        public void Initialize(IPluginHost host)
        {
            this.host = host;
            host.Log($"{Name} инициализирован");
        }

        public void Shutdown()
        {
            host?.Log($"{Name} выгружен");
        }

        public async Task ExportAsync(SearchResults results, string outputPath)
        {
            var html = new StringBuilder();
            html.AppendLine("<!DOCTYPE html>");
            html.AppendLine("<html><head><meta charset='utf-8'>");
            html.AppendLine("<title>Отчет FileCollector</title>");
            html.AppendLine("<style>body{font-family:Arial;margin:20px}table{border-collapse:collapse;width:100%}th,td{border:1px solid #ddd;padding:8px;text-align:left}th{background:#4CAF50;color:white}.stats{background:#f0f0f0;padding:15px;margin:10px 0;border-radius:5px}</style>");
            html.AppendLine("</head><body>");
            html.AppendLine($"<h1>Отчет FileCollector</h1>");
            html.AppendLine($"<div class='stats'>");
            html.AppendLine($"<p><b>Дата:</b> {results.SearchDate:dd.MM.yyyy HH:mm:ss}</p>");
            html.AppendLine($"<p><b>Найдено:</b> {results.TotalFound} | <b>Не найдено:</b> {results.NotFound}</p>");
            html.AppendLine("</div>");
            html.AppendLine("<table><tr><th>Имя файла</th><th>Расширение</th><th>Путь</th><th>Дата</th><th>Размер</th></tr>");
            
            foreach (var file in results.Files)
            {
                html.AppendLine($"<tr><td>{file.FileName}</td><td>{file.Extension}</td><td>{file.SourcePath}</td><td>{file.FileDate:dd.MM.yyyy HH:mm}</td><td>{FormatSize(file.FileSize)}</td></tr>");
            }
            
            html.AppendLine("</table></body></html>");
            
            await File.WriteAllTextAsync(outputPath, html.ToString(), Encoding.UTF8);
            host?.Log($"HTML отчет сохранен: {outputPath}");
        }

        private string FormatSize(long bytes)
        {
            if (bytes < 1024) return $"{bytes} B";
            if (bytes < 1024 * 1024) return $"{bytes / 1024.0:F2} KB";
            return $"{bytes / 1024.0 / 1024.0:F2} MB";
        }
    }
}
