using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using ClosedXML.Excel;
using Microsoft.Office.Interop.Excel;

namespace ZsuTools.Tables
{
    public class StroyovaZapyska
    {
        /// <summary>
        /// Дата, зазначена у Стройовій Записці.
        /// </summary>
        public DateTime Date { get; private set; }

        /// <summary>
        /// Військовослужбовці та де вони присутні в даний день. Також містить кольор запису
        /// </summary>
        public List<Entry> Items { get; private set; }

        /// <summary>
        /// Створює Стройову Записку з відкритої книги Excel (Interop). Дані читаються з першого аркуша.
        /// </summary>
        public StroyovaZapyska(Workbook workbook)
        {
            if (workbook == null || workbook.Worksheets.Count == 0)
            {
                throw new ArgumentException("Передано некоректну книгу Excel.", nameof(workbook));
            }

            // First worksheet is always Стройова Записка
            var worksheet = workbook.Worksheets[1] as Worksheet;

            Date = ParseDate(worksheet);
            Items = ParseMilitaryData(worksheet);
        }

        /// <summary>
        /// Створює Стройову Записку з Excel файлу, використовуючи ClosedXML (без Excel Interop).
        /// </summary>
        public StroyovaZapyska(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                throw new ArgumentException("Передано некоректний шлях до файлу.", nameof(filePath));
            }

            using (var workbook = new XLWorkbook(filePath))
            {
                Load(workbook);
            }
        }

        /// <summary>
        /// Створює Стройову Записку з Stream Excel файлу, використовуючи ClosedXML (без Excel Interop).
        /// </summary>
        public StroyovaZapyska(Stream stream)
        {
            if (stream == null)
            {
                throw new ArgumentException("Передано некоректний потік даних.", nameof(stream));
            }

            using (var workbook = new XLWorkbook(stream))
            {
                Load(workbook);
            }
        }

        private void Load(XLWorkbook workbook)
        {
            if (workbook.Worksheets.Count == 0)
            {
                throw new ArgumentException("Передано некоректну книгу Excel.");
            }

            // First worksheet is always Стройова Записка (ClosedXML worksheets are 1-based)
            var worksheet = workbook.Worksheet(1);
            Date = ParseDate(worksheet);
            Items = ParseMilitaryData(worksheet);
        }

        #region Parsing

        private const string DateCellAddress = "H3";

        #region Parsing: Excel Interop (COM)

        private static List<Entry> ParseMilitaryData(Worksheet worksheet)
        {
            var result = new List<Entry>();

            // Отримуємо використаний діапазон, щоб знайти останній рядок
            Range usedRange = worksheet.UsedRange;
            int lastRow = usedRange.Rows.Count + usedRange.Row - 1;

            int currentRow = 1;
            bool firstBlockFound = false;
            bool secondBlockFound = false;

            // --- КРОК 1: Шукаємо початок першого блоку (.NET Framework 4.8 сумісний) ---
            for (; currentRow <= lastRow; currentRow++)
            {
                var cellA = worksheet.Cells[currentRow, 1] as Range;
                var cellValue = GetMergedCellValue(cellA);

                if (cellValue.Contains("детально (відсутні)", StringComparison.InvariantCultureIgnoreCase))
                {
                    firstBlockFound = true;
                    currentRow += 2; // Пропускаємо рядок з написом + 1 рядок за умовою
                    break;
                }
            }

            if (!firstBlockFound)
            {
                return result;
            }

            // --- КРОК 2: Шукаємо, де починається другий блок ---
            int secondBlockStartRow = -1;
            for (int r = currentRow; r <= lastRow; r++)
            {
                var cellA = worksheet.Cells[r, 1] as Range;
                var cellValue = GetMergedCellValue(cellA);

                if (cellValue.Contains("бойове чергування", StringComparison.InvariantCultureIgnoreCase))
                {
                    secondBlockStartRow = r;
                    secondBlockFound = true;
                    break;
                }
            }

            // Обробляємо перший блок (до початку другого або до кінця листа)
            int firstBlockEndRow = secondBlockFound ? secondBlockStartRow - 1 : lastRow;
            for (; currentRow <= firstBlockEndRow; currentRow++)
            {
                Range cellA = worksheet.Cells[currentRow, 1] as Range;
                string cellValue = GetMergedCellValue(cellA);
                ProcessRow(worksheet, currentRow, new Place(LocationType.Відсутній, cellValue), result);
            }

            // --- КРОК 3: Обробка другого та третього блоків ---
            if (secondBlockFound)
            {
                currentRow = secondBlockStartRow + 1; // Починаємо одразу після рядка "бойове чергування"
                LocationType
                    currentLocation = LocationType.БЧ; // Початковий тип місцезнаходження після заголовка — другий

                for (; currentRow <= lastRow; currentRow++)
                {
                    Range cellA = worksheet.Cells[currentRow, 1] as Range;
                    string cellValue = GetMergedCellValue(cellA);

                    // Перевірка на початок ТРЕТЬОГО блоку (значення в колонці A рівне "ТИЛОВЕ ЗАБЕЗПЕЧЕННЯ")
                    if (cellValue.Equals("ТИЛОВЕ ЗАБЕЗПЕЧЕННЯ", StringComparison.OrdinalIgnoreCase))
                    {
                        currentLocation = LocationType.РЗ;
                    }

                    // Перевірка на кінець (тепер слово "Разом" закриває третій блок,
                    // але про всяк випадок зупинить цикл, навіть якщо третій блок чомусь не почався)
                    if (cellValue.Equals("Разом", StringComparison.OrdinalIgnoreCase))
                    {
                        break;
                    }

                    // Обробляємо рядок з поточним номером блоку (2 або 3)
                    ProcessRow(worksheet, currentRow, new Place(currentLocation, cellValue), result);
                }
            }

            return result;
        }

        /// <summary>
        /// Читає обов'язкову дату Стройової Записки з комірки H3 (COM).
        /// </summary>
        private static DateTime ParseDate(Worksheet worksheet)
        {
            Range cell = worksheet.Range[DateCellAddress];
            if (cell.MergeCells is bool merged && merged)
            {
                cell = cell.MergeArea.Cells[1, 1];
            }

            // Дата в Excel - це число (OADate)
            if (cell.Value2 is double oaDate)
            {
                return DateTime.FromOADate(oaDate);
            }

            string actualText = Convert.ToString(cell.Text);
            throw DateNotFound(actualText);
        }

        /// <summary>
        /// Обробляє окремий рядок, витягує людей з колонок C, D, E та шукає місцезнаходження в колонці A
        /// </summary>
        private static void ProcessRow(Worksheet worksheet, int row, Place location,
            List<Entry> result)
        {
            Range cellA = worksheet.Cells[row, 1] as Range;
            string locationStr = GetMergedCellValue(cellA);

            // Перевіряємо колонки C (3), D (4), E (5)
            for (int col = 3; col <= 5; col++)
            {
                Range personCell = worksheet.Cells[row, col] as Range;
                var personData = new RankPerson(personCell?.Value2?.ToString()?.Trim());

                if (!string.IsNullOrEmpty(personData.FullName))
                {
                    var cellColor = Utils.GetCellFontColor(personCell);
                    result.Add(new Entry(personData, location, cellColor));
                }
            }
        }

        /// <summary>
        /// Безпечно повертає значення комірки, враховуючи, чи вона є частиною об'єднаного діапазону
        /// </summary>
        private static string GetMergedCellValue(Range cell)
        {
            if (cell == null) return string.Empty;

            if (cell.MergeCells)
            {
                Range mergeArea = cell.MergeArea;
                Range topLeftCell = mergeArea.Cells[1, 1] as Range;
                return topLeftCell?.Value2?.ToString()?.Trim() ?? string.Empty;
            }

            return cell.Value2?.ToString()?.Trim() ?? string.Empty;
        }

        #endregion

        #region Parsing: ClosedXML

        /// <summary>
        /// Версія парсера на ClosedXML (без Excel Interop). Логіка ідентична COM-версії.
        /// </summary>
        private static List<Entry> ParseMilitaryData(IXLWorksheet worksheet)
        {
            var result = new List<Entry>();

            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
            int currentRow = 1;
            bool firstBlockFound = false;
            bool secondBlockFound = false;

            for (; currentRow <= lastRow; currentRow++)
            {
                var cellValue = GetMergedCellValue(worksheet.Cell(currentRow, 1));
                if (cellValue.Contains("детально (відсутні)", StringComparison.InvariantCultureIgnoreCase))
                {
                    firstBlockFound = true;
                    currentRow += 2;
                    break;
                }
            }

            if (!firstBlockFound)
            {
                return result;
            }

            int secondBlockStartRow = -1;
            for (int r = currentRow; r <= lastRow; r++)
            {
                var cellValue = GetMergedCellValue(worksheet.Cell(r, 1));
                if (cellValue.Contains("бойове чергування", StringComparison.InvariantCultureIgnoreCase))
                {
                    secondBlockStartRow = r;
                    secondBlockFound = true;
                    break;
                }
            }

            int firstBlockEndRow = secondBlockFound ? secondBlockStartRow - 1 : lastRow;
            for (; currentRow <= firstBlockEndRow; currentRow++)
            {
                string cellValue = GetMergedCellValue(worksheet.Cell(currentRow, 1));
                ProcessRow(worksheet, currentRow, new Place(LocationType.Відсутній, cellValue), result);
            }

            if (secondBlockFound)
            {
                currentRow = secondBlockStartRow + 1;
                LocationType currentLocation = LocationType.БЧ;

                for (; currentRow <= lastRow; currentRow++)
                {
                    string cellValue = GetMergedCellValue(worksheet.Cell(currentRow, 1));

                    if (cellValue.Equals("ТИЛОВЕ ЗАБЕЗПЕЧЕННЯ", StringComparison.OrdinalIgnoreCase))
                    {
                        currentLocation = LocationType.РЗ;
                    }

                    if (cellValue.Equals("Разом", StringComparison.OrdinalIgnoreCase))
                    {
                        break;
                    }

                    ProcessRow(worksheet, currentRow, new Place(currentLocation, cellValue), result);
                }
            }

            return result;
        }

        /// <summary>
        /// Читає обов'язкову дату Стройової Записки з комірки H3 (ClosedXML).
        /// </summary>
        private static DateTime ParseDate(IXLWorksheet worksheet)
        {
            var cell = worksheet.Cell(DateCellAddress);
            if (cell.IsMerged())
            {
                cell = cell.MergedRange().FirstCell();
            }

            if (cell.DataType == XLDataType.DateTime)
            {
                return cell.GetDateTime();
            }

            throw DateNotFound(cell.GetString());
        }

        private static void ProcessRow(IXLWorksheet worksheet, int row, Place location,
            List<Entry> result)
        {
            for (int col = 3; col <= 5; col++)
            {
                var personCell = worksheet.Cell(row, col);
                var personData = new RankPerson(personCell.GetString().Trim());

                if (!string.IsNullOrEmpty(personData.FullName))
                {
                    result.Add(new Entry(personData, location, Utils.GetCellFontColor(personCell)));
                }
            }
        }

        private static string GetMergedCellValue(IXLCell cell)
        {
            if (cell == null) return string.Empty;

            if (cell.IsMerged())
            {
                cell = cell.MergedRange().FirstCell();
            }

            return cell.GetString().Trim();
        }

        #endregion

        private static FormatException DateNotFound(string actualText)
        {
            return new FormatException(
                $"Комірка {DateCellAddress} Стройової Записки не містить дати (формат Date). Поточний вміст: '{actualText}'.");
        }

        #endregion

        /// <summary>
        /// Запис Стройової Записки: військовослужбовець, його стан (місцезнаходження) та колір запису.
        /// </summary>
        public sealed class Entry
        {
            public RankPerson Person { get; }
            public Place State { get; }
            public Color Color { get; }

            public Entry(RankPerson person, Place state, Color color)
            {
                Person = person;
                State = state;
                Color = color;
            }
        }

        public readonly struct Place
        {
            public readonly LocationType Location;
            public readonly string LocationName;

            public Place(LocationType location, string locationName)
            {
                Location = location;
                LocationName = locationName;
            }
        }

        public enum LocationType
        {
            Відсутній,
            БЧ,
            РЗ
        }
    }
}
