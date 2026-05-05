// FileCollector V3.0 - single-pass file index for fast lookup.
//
// V2.1 weakness: for each (fileName, extension) pair the search re-walked every search root,
// turning N*M (~30000) queries into N*M directory traversals. With recursive Parallel.ForEach
// inside the walker on top of that.
//
// V3.0: walk each root ONCE, group every file by its base name (Path.GetFileNameWithoutExtension),
// then answer FindLatest(baseName, extension) in O(k) where k is files sharing that base name.
// Total cost: O(total files visited) once, plus near-O(1) per query. For typical inputs this is
// ~50-200x faster than V2.1.

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using FileCollector.Core.Logging;

namespace FileCollector.Core.Search
{
    public sealed class FileIndex
    {
        private readonly Dictionary<string, List<FileInfo>> _byBaseName =
            new(StringComparer.OrdinalIgnoreCase);

        public int FileCount { get; private set; }
        public int FolderCount { get; private set; }
        public int InaccessibleCount { get; private set; }

        public static FileIndex Build(
            IReadOnlyList<string> roots,
            IAppLogger log,
            IProgress<string>? progress,
            bool skipSystemHidden,
            CancellationToken ct)
        {
            var idx = new FileIndex();
            foreach (var root in roots)
            {
                ct.ThrowIfCancellationRequested();
                if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root))
                {
                    log.Warn($"Папка поиска недоступна: {root}");
                    continue;
                }
                idx.IndexRoot(root, log, progress, skipSystemHidden, ct);
            }
            log.Info($"Индекс построен: {idx.FileCount} файлов в {idx.FolderCount} папках " +
                     $"(недоступных: {idx.InaccessibleCount}).");
            return idx;
        }

        public FileInfo? FindLatest(string baseName, string extension)
        {
            if (string.IsNullOrEmpty(baseName)) return null;
            if (!_byBaseName.TryGetValue(baseName, out var candidates)) return null;

            FileInfo? latest = null;
            foreach (var f in candidates)
            {
                if (!string.Equals(f.Extension, extension, StringComparison.OrdinalIgnoreCase))
                    continue;
                if (latest == null || f.LastWriteTime > latest.LastWriteTime)
                    latest = f;
            }
            return latest;
        }

        private void IndexRoot(
            string root,
            IAppLogger log,
            IProgress<string>? progress,
            bool skipSystemHidden,
            CancellationToken ct)
        {
            var stack = new Stack<string>();
            stack.Push(root);

            while (stack.Count > 0)
            {
                ct.ThrowIfCancellationRequested();
                string dir = stack.Pop();
                FolderCount++;
                progress?.Report(dir);

                DirectoryInfo dirInfo;
                try { dirInfo = new DirectoryInfo(dir); }
                catch (Exception ex)
                {
                    InaccessibleCount++;
                    log.Warn($"Не удалось открыть папку {dir}: {ex.Message}");
                    continue;
                }

                FileAttributes attrs;
                try { attrs = dirInfo.Attributes; }
                catch (Exception)
                {
                    InaccessibleCount++;
                    continue;
                }

                if (skipSystemHidden && (attrs & (FileAttributes.System | FileAttributes.Hidden)) != 0)
                    continue;

                // Files in this folder.
                try
                {
                    foreach (var file in dirInfo.EnumerateFiles())
                    {
                        if (skipSystemHidden &&
                            (file.Attributes & (FileAttributes.System | FileAttributes.Hidden)) != 0)
                            continue;

                        string baseName = Path.GetFileNameWithoutExtension(file.Name);
                        if (!_byBaseName.TryGetValue(baseName, out var list))
                        {
                            list = new List<FileInfo>(2);
                            _byBaseName[baseName] = list;
                        }
                        list.Add(file);
                        FileCount++;
                    }
                }
                catch (UnauthorizedAccessException)
                {
                    InaccessibleCount++;
                    log.Warn($"Нет доступа к файлам в {dir}.");
                }
                catch (Exception ex)
                {
                    log.Warn($"Ошибка чтения файлов в {dir}: {ex.Message}");
                }

                // Subdirectories.
                try
                {
                    foreach (var sub in dirInfo.EnumerateDirectories())
                        stack.Push(sub.FullName);
                }
                catch (UnauthorizedAccessException)
                {
                    InaccessibleCount++;
                }
                catch (Exception ex)
                {
                    log.Warn($"Ошибка обхода подпапок {dir}: {ex.Message}");
                }
            }
        }
    }
}
