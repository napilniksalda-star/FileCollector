// FileCollector V3.0 - path-traversal protection (new in V3.0; V2.1 had no validation).
//
// V2.1 issue: destination path was Path.Combine(destinationPath, Path.GetFileName(sourcePath)).
// If a future caller ever passed an unsanitized name, that would still allow traversal because
// Path.GetFileName preserves "..". V3.0 explicitly verifies the resolved combined path lives
// under the destination root and rejects names that contain invalid characters.

using System;
using System.IO;

namespace FileCollector.Core.Copy
{
    public static class PathSafety
    {
        /// <summary>
        /// Combine <paramref name="baseDir"/> and <paramref name="candidateFileName"/> safely.
        /// Returns the absolute combined path if it lives under baseDir; null otherwise.
        /// </summary>
        public static string? SafeCombine(string baseDir, string candidateFileName)
        {
            if (string.IsNullOrWhiteSpace(baseDir) || string.IsNullOrWhiteSpace(candidateFileName))
                return null;

            // Strip any directory components from the candidate; prevents "..\evil.exe".
            string safeName = Path.GetFileName(candidateFileName);
            if (string.IsNullOrEmpty(safeName)) return null;
            if (safeName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) return null;

            string baseFull;
            try { baseFull = Path.GetFullPath(baseDir); }
            catch { return null; }
            baseFull = baseFull.TrimEnd('\\', '/');

            string combinedFull;
            try { combinedFull = Path.GetFullPath(Path.Combine(baseFull, safeName)); }
            catch { return null; }

            string baseWithSep = baseFull + Path.DirectorySeparatorChar;
            if (!combinedFull.StartsWith(baseWithSep, StringComparison.OrdinalIgnoreCase))
                return null;

            return combinedFull;
        }

        public static bool PathsEqual(string a, string b)
        {
            try
            {
                return string.Equals(Path.GetFullPath(a), Path.GetFullPath(b),
                    StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }
    }
}
