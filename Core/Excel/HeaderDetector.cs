// FileCollector V3.0 - heuristic detection of header / total / placeholder cells.
// Behavior preserved from V2.1 IsHeaderOrException, just relocated and cleaned up.

using System;
using System.Collections.Generic;
using System.Linq;

namespace FileCollector.Core.Excel
{
    public static class HeaderDetector
    {
        private static readonly HashSet<string> KnownHeaders = new(StringComparer.OrdinalIgnoreCase)
        {
            // Russian
            "Имя файла", "File Name", "Наименование файла", "Название файла", "Имя", "Наименование",
            "Название", "Файл", "Файлы", "Документ", "Документы", "Чертеж", "Чертежи", "Модель", "Модели",
            "Схема", "Схемы", "Проект", "Проекты", "Архив", "Архивы", "Папка", "Папки",
            // English
            "Filename", "File", "Name", "Document", "Drawing", "Model", "Scheme", "Project", "Folder",
            "File name", "File names", "Document name", "Drawing name",
            // Table-position headers
            "№ п/п", "№", "Порядковый номер", "Номер", "Позиция", "Код", "Артикул", "Артикулы",
            "Количество", "Кол-во", "Кол.", "Примечание", "Примечания", "Комментарий", "Комментарии",
            "Дата", "Date", "Время", "Time", "Создан", "Создано", "Изменен", "Изменено",
            "Автор", "Author", "Версия", "Version", "Ревизия", "Revision", "Rev.", "Вер.",
            // Totals
            "Итого", "Всего", "Total", "Сумма", "Sum", "Итог", "Итоги", "Total sum",
            "Всего строк", "Всего записей", "Total records", "Итого по странице",
            "Конец", "End", "Конец списка", "End of list", "Продолжение", "Continuation",
            "Далее", "Next", "Следующая страница", "Next page",
            // Placeholders
            "н/д", "n/a", "N/A", "не указано", "не задано", "не применимо",
            "пусто", "empty", "null", "undefined", "отсутствует", "missing",
            // Symbols
            "-", "--", "---", "...", "…", "***", "///"
        };

        private static readonly string[] HeaderPrefixes =
        {
            "Итого", "Всего", "Total", "Sum"
        };

        private static readonly string[] PageMarkerPrefixes =
        {
            "страница", "page", "таблица", "table", "отчет", "report"
        };

        public static bool IsHeaderOrException(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return false;

            string trimmed = value.Trim();

            if (KnownHeaders.Contains(trimmed)) return true;

            // "Header:" / "Header -" / "Header (..."
            foreach (var h in KnownHeaders)
            {
                if (!trimmed.StartsWith(h, StringComparison.OrdinalIgnoreCase)) continue;
                if (trimmed.Length == h.Length) return true;
                char next = trimmed[h.Length];
                if (next == ':' || next == ' ' || next == '-' || next == '(') return true;
            }

            if (HeaderPrefixes.Any(p => trimmed.StartsWith(p, StringComparison.OrdinalIgnoreCase)))
                return true;

            // Pure number
            if (int.TryParse(trimmed, out _)) return true;

            // Pure date
            if (DateTime.TryParse(trimmed, out _)) return true;

            // Page/table/report markers (only if at the beginning)
            string lower = trimmed.ToLowerInvariant();
            if (PageMarkerPrefixes.Any(p => lower.StartsWith(p))) return true;

            return false;
        }
    }
}
