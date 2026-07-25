using System;
using System.Collections.Generic;
using System.Drawing;
using Microsoft.Office.Interop.Excel;

namespace ZsuTools
{
    public class SZParser
    {
        public static List<Tuple<RankPerson, SZState, Color>> ParseMilitaryData(Worksheet worksheet)
        {
            var result = new List<Tuple<RankPerson, SZState, Color>>();

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
                ProcessRow(worksheet, currentRow, new SZState(LocationType.Відсутній, cellValue), result);
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
                    ProcessRow(worksheet, currentRow, new SZState(currentLocation, cellValue), result);
                }
            }

            return result;
        }

        /// <summary>
        /// Обробляє окремий рядок, витягує людей з колонок C, D, E та шукає місцезнаходження в колонці A
        /// </summary>
        private static void ProcessRow(Worksheet worksheet, int row, SZState location,
            List<Tuple<RankPerson, SZState, Color>> result)
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
                    result.Add(Tuple.Create(personData, location, cellColor));
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
    }

    public readonly struct SZState
    {
        public readonly LocationType Location;
        public readonly string LocationName;

        public SZState(LocationType location, string locationName)
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