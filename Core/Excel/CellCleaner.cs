// FileCollector V3.0 - cell-text normalizer (extracted from MainForm.GetCellValueAsString).
// Strips invisible Unicode space variants, smart quotes, dashes, and BOMs that creep into copy-pasted Excel data.
// All non-ASCII characters declared as \uXXXX escapes so the file's encoding cannot lose them.

using System;
using System.Text;
using OfficeOpenXml;

namespace FileCollector.Core.Excel
{
    public static class CellCleaner
    {
        // Whitespace variants - collapsed to ASCII space.
        private static readonly char[] SpaceLikeChars =
        {
            ' ', // NBSP
            ' ', // narrow NBSP
            ' ', // en quad
            ' ', // em quad
            ' ', // en space
            ' ', // em space
            ' ', // three-per-em
            ' ', // four-per-em
            ' ', // six-per-em
            ' ', // figure space
            ' ', // punctuation space
            ' ', // thin space
            ' ', // hair space
            '\r', '\n', '\t'
        };

        // Quote variants - erased.
        private static readonly char[] QuoteChars =
        {
            '"', '\'',
            '“', '”', // " "
            '‘', '’', // ' '
            '«', '»', // <<  >>
            '‹', '›', // single-angle quotes
            '`',
            '´'            // acute accent
        };

        // Dash variants - normalized to ASCII '-'.
        private static readonly char[] DashChars =
        {
            '–', // en dash
            '—', // em dash
            '‐', // hyphen
            '‑', // non-breaking hyphen
            '⁃', // hyphen bullet
            '−', // minus
            '─'  // box-drawings horizontal
        };

        // Invisible characters - erased.
        private static readonly char[] InvisibleChars =
        {
            '​', // zero-width space
            '‌', // zero-width non-joiner
            '‍', // zero-width joiner
            '⁫', // activate symmetric swapping
            '﻿'  // zero-width no-break space / BOM
        };

        public static string Clean(ExcelRange? cell)
        {
            if (cell?.Value == null) return string.Empty;

            string text;
            try { text = cell.Text; }
            catch { text = cell.Value?.ToString() ?? string.Empty; }
            if (string.IsNullOrEmpty(text))
                text = cell.Value?.ToString() ?? string.Empty;
            if (string.IsNullOrEmpty(text)) return string.Empty;

            return Clean(text);
        }

        public static string Clean(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;

            var buf = new StringBuilder(text.Length);
            foreach (var ch in text)
            {
                if (Array.IndexOf(InvisibleChars, ch) >= 0) continue;
                if (Array.IndexOf(SpaceLikeChars, ch) >= 0) { buf.Append(' '); continue; }
                if (Array.IndexOf(QuoteChars, ch) >= 0) continue;
                if (Array.IndexOf(DashChars, ch) >= 0) { buf.Append('-'); continue; }
                buf.Append(ch);
            }

            string s = buf.ToString();
            while (s.Contains("  ")) s = s.Replace("  ", " ");
            return s.Trim();
        }
    }
}
