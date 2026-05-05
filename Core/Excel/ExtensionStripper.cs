// FileCollector V3.0 - removes a known file extension from the end of a name (case-insensitive).
// Mirrors V2.1 GetFileNameWithoutExtension behavior - only strips extensions we know about,
// so a name like "Bracket M16" doesn't accidentally lose the "M16" suffix.

using System;
using System.Collections.Generic;

namespace FileCollector.Core.Excel
{
    public static class ExtensionStripper
    {
        private static readonly HashSet<string> KnownExtensions = new(StringComparer.OrdinalIgnoreCase)
        {
            ".pdf", ".dxf", ".dwg", ".doc", ".docx", ".xlsx", ".xls", ".txt",
            ".jpg", ".png", ".sldprt", ".sldasm", ".slddrw", ".step", ".stp",
            ".iges", ".igs", ".ipt", ".iam", ".idw", ".prt", ".asm", ".drw",
            ".catpart", ".catproduct", ".catdrawing", ".par", ".psm", ".dft",
            ".3dm", ".skp", ".dgn", ".rvt", ".rfa", ".rte", ".ifc", ".sat",
            ".x_t", ".x_b", ".jt", ".u3d", ".dae", ".fbx", ".obj", ".stl"
        };

        public static string Strip(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName)) return fileName;

            foreach (var ext in KnownExtensions)
            {
                if (fileName.EndsWith(ext, StringComparison.OrdinalIgnoreCase))
                    return fileName.Substring(0, fileName.Length - ext.Length).Trim();
            }
            return fileName;
        }
    }
}
