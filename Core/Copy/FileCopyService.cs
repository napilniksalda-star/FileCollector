// FileCollector V3.0 - async file copy with retry, locked-file detection, path safety.
//
// V2.1 issues fixed:
//   - Thread.Sleep(1000) blocked the worker; CancellationToken couldn't break the wait. V3.0 uses Task.Delay(..., ct).
//   - File.Copy() is synchronous; V3.0 uses async stream copy with FileOptions.Asynchronous so the OS scheduler
//     can interleave the copy with other I/O.
//   - String-based status field replaced by CopyStatus enum.
//   - Destination path validated via PathSafety (path-traversal guard).

using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using FileCollector.Core.Logging;

namespace FileCollector.Core.Copy
{
    public sealed class FileCopyService
    {
        private const int BufferSize = 64 * 1024;
        private const int MaxAttempts = 3;
        private const int RetryDelayMs = 1000;

        private readonly IAppLogger _log;

        public FileCopyService(IAppLogger log) { _log = log; }

        public async Task<CopyOperation> CopyAsync(
            FileInfo source,
            string destinationDir,
            string fileName,
            string extension,
            bool skipLocked,
            CancellationToken ct)
        {
            var op = new CopyOperation
            {
                FileName = fileName,
                Extension = extension,
                SourcePath = source.FullName
            };

            if (!source.Exists)
            {
                op.Status = CopyStatus.NotFound;
                op.Message = "Исходный файл больше не существует";
                return op;
            }

            string? safeDest = PathSafety.SafeCombine(destinationDir, source.Name);
            if (safeDest is null)
            {
                op.Status = CopyStatus.PathInvalid;
                op.Message = $"Недопустимое имя файла или путь: {source.Name}";
                _log.Warn($"Отклонён небезопасный путь назначения для '{source.Name}'.");
                return op;
            }
            op.DestinationPath = safeDest;

            if (PathSafety.PathsEqual(source.FullName, safeDest))
            {
                op.Status = CopyStatus.AlreadyAtDestination;
                op.Message = "Файл уже находится в папке назначения";
                op.FileSize = source.Length;
                return op;
            }

            for (int attempt = 1; attempt <= MaxAttempts; attempt++)
            {
                ct.ThrowIfCancellationRequested();
                try
                {
                    await using var src = new FileStream(
                        source.FullName, FileMode.Open, FileAccess.Read, FileShare.Read,
                        BufferSize, FileOptions.Asynchronous | FileOptions.SequentialScan);
                    await using var dst = new FileStream(
                        safeDest, FileMode.Create, FileAccess.Write, FileShare.None,
                        BufferSize, FileOptions.Asynchronous);
                    await src.CopyToAsync(dst, ct).ConfigureAwait(false);

                    op.Status = CopyStatus.Success;
                    op.FileSize = source.Length;
                    return op;
                }
                catch (OperationCanceledException)
                {
                    TryDeletePartial(safeDest);
                    throw;
                }
                catch (IOException ex) when (IsLockedFile(ex))
                {
                    if (skipLocked)
                    {
                        op.Status = CopyStatus.Locked;
                        op.Message = "Файл заблокирован, пропуск";
                        return op;
                    }
                    if (attempt < MaxAttempts)
                    {
                        _log.Info($"Попытка {attempt}: файл занят, повтор через 1 сек: {source.Name}");
                        await Task.Delay(RetryDelayMs, ct).ConfigureAwait(false);
                        continue;
                    }
                    op.Status = CopyStatus.Locked;
                    op.Message = "Файл заблокирован после " + MaxAttempts + " попыток";
                    return op;
                }
                catch (UnauthorizedAccessException ex)
                {
                    op.Status = CopyStatus.AccessDenied;
                    op.Message = ex.Message;
                    return op;
                }
                catch (Exception ex)
                {
                    op.Status = CopyStatus.Failed;
                    op.Message = ex.Message;
                    _log.Error($"Ошибка копирования {source.FullName}: {ex.Message}");
                    return op;
                }
            }

            op.Status = CopyStatus.Failed;
            op.Message = "Не удалось скопировать после повторных попыток";
            return op;
        }

        private static bool IsLockedFile(IOException ex)
        {
            int code = Marshal.GetHRForException(ex) & 0xFFFF;
            return code == 32 || code == 33; // ERROR_SHARING_VIOLATION, ERROR_LOCK_VIOLATION
        }

        private static void TryDeletePartial(string path)
        {
            try { if (File.Exists(path)) File.Delete(path); }
            catch { /* best-effort cleanup */ }
        }
    }
}
