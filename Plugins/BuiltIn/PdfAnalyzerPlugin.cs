using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace FileCollector.Plugins.BuiltIn
{
    public class PdfAnalyzerPlugin : IFileProcessorPlugin
    {
        private IPluginHost? host;

        public string Name => "PDF Analyzer";
        public string Version => "1.0";
        public string Description => "Анализирует PDF файлы: размер, количество страниц";

        public void Initialize(IPluginHost host)
        {
            this.host = host;
            host.Log($"{Name} инициализирован");
        }

        public void Shutdown()
        {
            host?.Log($"{Name} выгружен");
        }

        public bool CanProcess(string extension) => 
            extension.Equals(".pdf", StringComparison.OrdinalIgnoreCase);

        public async Task<ProcessResult> ProcessAsync(FileInfo file, CancellationToken ct)
        {
            await Task.Delay(10, ct);
            
            var result = new ProcessResult { Success = true };
            result.Metadata["FileSize"] = file.Length;
            result.Metadata["FileSizeMB"] = Math.Round(file.Length / 1024.0 / 1024.0, 2);
            result.Message = $"PDF: {result.Metadata["FileSizeMB"]} MB";
            
            return result;
        }
    }
}
