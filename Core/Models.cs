// FileCollector V3.0 - shared model types (extracted from MainForm.cs in V2.1).

using System;

namespace FileCollector.Core
{
    /// <summary>
    /// One row in the preview grid: a (fileName, extension) pair plus the file we found for it (if any).
    /// V3.0: SourcePath is null when nothing was found (V2.1 used empty string - ambiguous).
    /// </summary>
    public sealed class PreviewItem
    {
        public required string FileName { get; init; }
        public required string Extension { get; init; }
        public string? SourcePath { get; set; }
        public DateTime? FileDate { get; set; }

        public bool IsFound => !string.IsNullOrEmpty(SourcePath);
    }

    /// <summary>
    /// Outcome of one copy attempt. V3.0: replaces the V2.1 free-form "Status" string with a typed enum.
    /// </summary>
    public sealed class CopyOperation
    {
        public required string FileName { get; init; }
        public required string Extension { get; init; }
        public string SourcePath { get; init; } = string.Empty;
        public string DestinationPath { get; set; } = string.Empty;
        public CopyStatus Status { get; set; }
        public string? Message { get; set; }
        public long FileSize { get; set; }
    }

    public enum CopyStatus
    {
        NotFound,
        PathInvalid,
        AccessDenied,
        Locked,
        AlreadyAtDestination,
        Failed,
        Success
    }

    /// <summary>
    /// Aggregated counters surfaced to the UI.
    /// </summary>
    public sealed class RunSummary
    {
        public int Total { get; set; }
        public int Copied { get; set; }
        public int Skipped { get; set; }
        public int NotFound { get; set; }
        public int Failed { get; set; }
    }
}
