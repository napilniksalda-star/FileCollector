using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace FileCollector.Plugins
{
    public interface IPlugin
    {
        string Name { get; }
        string Version { get; }
        string Description { get; }
        void Initialize(IPluginHost host);
        void Shutdown();
    }

    public interface IPluginHost
    {
        void Log(string message);
        string GetSetting(string key);
    }

    public interface IFileProcessorPlugin : IPlugin
    {
        bool CanProcess(string extension);
        Task<ProcessResult> ProcessAsync(FileInfo file, CancellationToken ct);
    }

    public interface IExportPlugin : IPlugin
    {
        string FormatName { get; }
        Task ExportAsync(SearchResults results, string outputPath);
    }

    public class ProcessResult
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public Dictionary<string, object> Metadata { get; set; } = new();
    }

    public class SearchResults
    {
        public List<FileResult> Files { get; set; } = new();
        public int TotalFound { get; set; }
        public int NotFound { get; set; }
        public DateTime SearchDate { get; set; }
    }

    public class FileResult
    {
        public string FileName { get; set; } = string.Empty;
        public string Extension { get; set; } = string.Empty;
        public string SourcePath { get; set; } = string.Empty;
        public DateTime FileDate { get; set; }
        public long FileSize { get; set; }
    }
}
