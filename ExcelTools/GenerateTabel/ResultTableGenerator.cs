using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;

namespace ZsuTools
{
    public class ResultTableGenerator
    {
        public static void GenerateSummaryTable(Dictionary<int, List<Tuple<RankPerson, SZState, Color>>> monthlyResults)
        {
            if (monthlyResults == null || monthlyResults.Count == 0)
            {
                MessageBox.Show("Немає даних для формування таблиці.", "Повідомлення", MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            // Отримуємо доступ до Excel
            Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;

            bool previousScreenUpdating = excelApp.ScreenUpdating;
            excelApp.ScreenUpdating = false;

            // 1. Створюємо нову книгу з одним листом
            Excel.Workbook newWorkbook = excelApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            Excel.Worksheet sheet = (Excel.Worksheet)newWorkbook.Worksheets[1];
            sheet.Name = "Зведена відомість";

            // Визначаємо максимальну кількість днів у місяці на основі наявних ключів (наприклад, 30 або 31)
            int maxDay = monthlyResults.Keys.Max();

            // 2. Формуємо шапку (перший рядок)
            sheet.Cells[1, 1] = "Звання";
            sheet.Cells[1, 2] = "ПРІЗВИЩЕ (за наявності) Ім'я По батькові (за наявності)";

            for (int day = 1; day <= maxDay; day++)
            {
                // Дні починаються з 3-ї колонки (C)
                sheet.Cells[1, 2 + day] = day;
            }

            // Структура для відслідковування рядків: Ключ — ПІБ (унікальний ідентифікатор), Значення — номер рядка в Excel
            var personRowMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            int nextAvailableRow = 2; // Дані людей починаються з 2-го рядка

            // 3. Заповнюємо дані по днях
            // Сортуємо дні по порядку (від 1 до maxDay), щоб послідовно заповнювати колонки
            var sortedDays = monthlyResults.Keys.OrderBy(d => d).ToList();

            foreach (int day in sortedDays)
            {
                int targetColumn = 2 + day; // Номер колонки для поточного дня
                var dayRecords = monthlyResults[day];

                foreach (var record in dayRecords)
                {
                    var person = record.Item1; 
                    var location = record.Item2; 

                    // Розбиваємо на Звання та ПІБ
                    var rank = person.Rank;
                    var fullName = person.FullName;

                    int targetRow;

                    // Перевіряємо, чи ця людина вже є в нашому реєстрі (таблиці)
                    if (personRowMap.TryGetValue(fullName, out var valueRow))
                    {
                        // Якщо є, беремо її зафіксований рядок
                        targetRow = valueRow;
                    }
                    else
                    {
                        // Якщо людина зустрілася вперше в цьому місяці, виділяємо їй новий рядок
                        targetRow = nextAvailableRow;
                        personRowMap.Add(fullName, targetRow);

                        // Записуємо базові дані (Звання в кол. A, ПІБ в кол. B)
                        sheet.Cells[targetRow, 1] = rank;
                        sheet.Cells[targetRow, 2] = fullName;

                        nextAvailableRow++;
                    }

                    // Записуємо тип місцеположення у клітинку відповідного дня
                    sheet.Cells[targetRow, targetColumn] = LocationToTabelState(location.Location, location.LocationName);
                    //sheet.Cells[targetRow, targetColumn] = location.LocationName;
                    sheet.Cells[targetRow, targetColumn].Font.Color = ColorTranslator.ToOle(record.Item3);
                }
            }

            // 4. Гарне форматування отриманої таблиці
            FormatResultTable(sheet, maxDay, nextAvailableRow - 1);

            excelApp.ScreenUpdating = previousScreenUpdating;
        }

        private static string LocationToTabelState(LocationType location, string locationStr)
        {
            switch (location)
            {
                case LocationType.БЧ:
                    return "'++";
                case LocationType.РЗ:
                    return "+";
                case LocationType.Відсутній:
                    if (locationStr.StartsWith("Відрядження", StringComparison.InvariantCultureIgnoreCase)) return "вдр";
                    else if (locationStr.StartsWith("Відпустка", StringComparison.InvariantCultureIgnoreCase)) return "від";
                    else if (locationStr.StartsWith("Лікування", StringComparison.InvariantCultureIgnoreCase)) return "лік";
                    else if(locationStr.StartsWith("ВЛК", StringComparison.InvariantCultureIgnoreCase)) return "вдр";
                    else if(locationStr.StartsWith("СЗЧ", StringComparison.InvariantCultureIgnoreCase)) return "СЗЧ";
                    else if(locationStr.StartsWith("Зниклі безвісті", StringComparison.InvariantCultureIgnoreCase)) return "ЗБ";
                    else if (locationStr.StartsWith("Загиблі", StringComparison.InvariantCultureIgnoreCase))
                        return "заг";
                    else return locationStr;
                default:
                    return string.Empty;
            }
        }

        /// <summary>
        /// Наведення базового військового порядку та краси в Excel таблиці
        /// </summary>
        private static void FormatResultTable(Excel.Worksheet sheet, int maxDay, int totalRows)
        {
            // Оформлюємо шапку (жирний шрифт, вирівнювання по центру)
            Excel.Range headerRange = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, 2 + maxDay]];
            headerRange.Font.Bold = true;
            headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            // Межі (Borders) для всієї таблиці
            Excel.Range tableRange = sheet.Range[sheet.Cells[1, 1], sheet.Cells[totalRows, 2 + maxDay]];
            tableRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            tableRange.Borders.Weight = Excel.XlBorderWeight.xlThin;
            //Set font and size
            tableRange.Font.Name = "Times New Roman";
            tableRange.Font.Size = 14;

            // Колонки з днями вирівнюємо по центру
            Excel.Range daysDataRange = sheet.Range[sheet.Cells[2, 3], sheet.Cells[totalRows, 2 + maxDay]];
            daysDataRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            // Автопідбір ширини для перших двох колонок (Звання та ПІБ)
            Excel.Range firstTwoColumns = sheet.Range["A:B"];
            firstTwoColumns.EntireColumn.AutoFit();

            // Для днів робимо фіксовану компактну ширину
            for (int day = 1; day <= maxDay; day++)
            {
                ((Excel.Range)sheet.Cells[1, 2 + day]).EntireColumn.ColumnWidth = 4;
            }
        }
    }
}