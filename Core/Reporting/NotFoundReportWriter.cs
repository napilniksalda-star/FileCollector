// FileCollector V3.0 - writes a grouped report of files we failed to find/copy.
// Behaviorally similar to V2.1, but stamps the V3.0 banner and triggers on a wider set of statuses.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using FileCollector.Core.Logging;

namespace FileCollector.Core.Reporting
{
    public sealed class NotFoundReportWriter
    {
        private static readonly CopyStatus[] FailureStatuses =
        {
            CopyStatus.NotFound,
            CopyStatus.PathInvalid,
            CopyStatus.AccessDenied,
            CopyStatus.Locked,
            CopyStatus.Failed
        };

        private readonly IAppLogger _log;
        public NotFoundReportWriter(IAppLogger log) { _log = log; }

        public string? Write(string destinationDir, IEnumerable<CopyOperation> operations)
        {
            if (string.IsNullOrEmpty(destinationDir) || !Directory.Exists(destinationDir))
                return null;

            var failures = operations.Where(op => Array.IndexOf(FailureStatuses, op.Status) >= 0).ToList();
            if (failures.Count == 0) return null;

            string ts = DateTime.Now.ToString("dd.MM.yy HH.mm");
            string path = Path.Combine(destinationDir, $"Отчёт не найденных файлов {ts}.txt");

            try
            {
                using var sw = new StreamWriter(path, false, new UTF8Encoding(true));
                sw.WriteLine("Отчёт о ненайденных и непрокопированных файлах");
                sw.WriteLine($"FileCollector V{VersionInfo.Version} ({VersionInfo.CodeName})");
                sw.WriteLine($"Дата создания: {DateTime.Now:dd.MM.yyyy HH:mm:ss}");
                sw.WriteLine($"Всего записей: {failures.Count}");
                sw.WriteLine(new string('=', 60));
                sw.WriteLine();

                foreach (var statusGroup in failures
                             .GroupBy(op => op.Status)
                             .OrderBy(g => g.Key))
                {
                    sw.WriteLine($"=== {DescribeStatus(statusGroup.Key)} ({statusGroup.Count()}) ===");
                    foreach (var extGroup in statusGroup
                                 .GroupBy(op => op.Extension)
                                 .OrderBy(g => g.Key))
                    {
                        sw.WriteLine($"  Формат: {extGroup.Key} (всего: {extGroup.Count()})");
                        sw.WriteLine("  " + new string('-', 38));
                        foreach (var op in extGroup.OrderBy(o => o.FileName))
                        {
                            string? note = string.IsNullOrEmpty(op.Message) ? null : $"  // {op.Message}";
                            sw.WriteLine($"    {op.FileName}{op.Extension}{note}");
                        }
                        sw.WriteLine();
                    }
                }

                _log.Info($"Создан отчёт: {path}");
                return path;
            }
            catch (Exception ex)
            {
                _log.Error($"Ошибка при создании отчёта: {ex.Message}");
                return null;
            }
        }

        private static string DescribeStatus(CopyStatus s) => s switch
        {
            CopyStatus.NotFound       => "Не найдено в папках поиска",
            CopyStatus.PathInvalid    => "Недопустимый путь / небезопасное имя",
            CopyStatus.AccessDenied   => "Нет прав доступа",
            CopyStatus.Locked         => "Файл заблокирован",
            CopyStatus.Failed         => "Ошибка копирования",
            _                         => s.ToString()
        };
    }
}
