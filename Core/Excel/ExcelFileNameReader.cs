// FileCollector V3.0 - Excel reader extracted from MainForm.cs.
// Reads a list of base file names from column B of the first worksheet.
// V3.0 changes vs V2.1:
//   - No UI dependencies (takes IAppLogger instead of calling LogMessage directly).
//   - Empty workbook / missing sheet returns empty list with a warning instead of throwing.
//   - Header detection rules unchanged - kept the V2.1 behavior intact, just relocated.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using FileCollector.Core.Logging;
using OfficeOpenXml;

namespace FileCollector.Core.Excel
{
    public sealed class ExcelFileNameReader
    {
        private readonly IAppLogger _log;

        public ExcelFileNameReader(IAppLogger log)
        {
            _log = log;
        }

        public List<string> Read(string excelPath)
        {
            var fileNames = new List<string>();

            try
            {
                using var package = new ExcelPackage(new FileInfo(excelPath));

                if (package.Workbook.Worksheets.Count == 0)
                {
                    _log.Warn("В Excel файле нет ни одного листа.");
                    return fileNames;
                }

                var ws = package.Workbook.Worksheets[0];
                if (ws.Dimension == null)
                {
                    _log.Warn("Excel файл пустой или не содержит данных.");
                    return fileNames;
                }

                int rowCount = ws.Dimension.End.Row;
                _log.Info($"Размер таблицы: {rowCount} строк.");

                const int fileNameColumn = 2; // column B
                int startRow = FindDataStartRow(ws, fileNameColumn, rowCount);
                _log.Info($"Чтение данных из столбца B, начиная со строки {startRow}.");

                int added = 0, skipped = 0;
                for (int row = startRow; row <= rowCount; row++)
                {
                    string raw = CellCleaner.Clean(ws.Cells[row, fileNameColumn]);
                    if (string.IsNullOrWhiteSpace(raw))
                    {
                        skipped++;
                        continue;
                    }

                    string trimmed = raw.Trim();
                    if (HeaderDetector.IsHeaderOrException(trimmed))
                    {
                        _log.Debug($"Пропущена ячейка B{row} (заголовок/служебная): '{trimmed}'");
                        skipped++;
                        continue;
                    }

                    string nameWithoutExt = ExtensionStripper.Strip(trimmed);
                    string finalName = string.IsNullOrWhiteSpace(nameWithoutExt) ? trimmed : nameWithoutExt;
                    fileNames.Add(finalName);
                    added++;
                }

                _log.Info($"Загружено {added} имён файлов из Excel (пропущено {skipped}).");

                if (added == 0)
                    _log.Warn("Не найдено ни одного имени файла в столбце B.");
            }
            catch (Exception ex)
            {
                _log.Error($"Ошибка чтения Excel: {ex.Message}");
            }

            return fileNames
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private int FindDataStartRow(ExcelWorksheet worksheet, int column, int rowCount)
        {
            int rowsToCheck = Math.Min(5, rowCount);
            for (int row = 1; row <= rowsToCheck; row++)
            {
                string val = CellCleaner.Clean(worksheet.Cells[row, column]).Trim();
                if (val.Length == 0) continue;
                if (!HeaderDetector.IsHeaderOrException(val))
                    return row;
                _log.Debug($"Строка B{row} - заголовок: '{val}', пропускаем.");
            }
            return 1;
        }
    }
}
