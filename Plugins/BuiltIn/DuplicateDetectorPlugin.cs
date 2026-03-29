using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;

namespace FileCollector.Plugins.BuiltIn
{
    public class DuplicateDetectorPlugin : IFileProcessorPlugin
    {
        private IPluginHost? host;
        private Dictionary<string, List<string>> hashMap = new();

        public string Name => "Duplicate Detector";
        public string Version => "1.0";
        public string Description => "Обнаружение дубликатов файлов по MD5 хешу";

        public void Initialize(IPluginHost host)
        {
            this.host = host;
            host.Log($"{Name} инициализирован");
        }

        public void Shutdown()
        {
            host?.Log($"{Name} выгружен - найдено {hashMap.Count(kvp => kvp.Value.Count > 1)} групп дубликатов");
            hashMap.Clear();
        }

        public bool CanProcess(string extension) => true;

        public async Task<ProcessResult> ProcessAsync(FileInfo file, CancellationToken ct)
        {
            var result = new ProcessResult { Success = true };
            
            try
            {
                string hash = await ComputeMD5Async(file.FullName, ct);
                
                if (!hashMap.ContainsKey(hash))
                    hashMap[hash] = new List<string>();
                
                hashMap[hash].Add(file.FullName);
                
                if (hashMap[hash].Count > 1)
                {
                    result.Message = $"Дубликат! Найдено {hashMap[hash].Count} копий";
                    result.Metadata["IsDuplicate"] = true;
                    result.Metadata["DuplicateCount"] = hashMap[hash].Count;
                }
                else
                {
                    result.Message = "Уникальный файл";
                    result.Metadata["IsDuplicate"] = false;
                }
                
                result.Metadata["Hash"] = hash;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Ошибка: {ex.Message}";
            }
            
            return result;
        }

        private async Task<string> ComputeMD5Async(string filePath, CancellationToken ct)
        {
            using var md5 = MD5.Create();
            using var stream = File.OpenRead(filePath);
            var hash = await md5.ComputeHashAsync(stream, ct);
            return BitConverter.ToString(hash).Replace("-", "").ToLower();
        }
    }
}
