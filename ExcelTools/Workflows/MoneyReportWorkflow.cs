using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Xceed.Words.NET;
using ZsuTools.Tables;
using ZsuTools.UI;
using Application = Microsoft.Office.Interop.Excel.Application;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using Word = Microsoft.Office.Interop.Word;
using XDoc = Xceed.Document.NET;

namespace ZsuTools.Workflows
{
    /// <summary>
    /// Процеси створення грошових рапортів на основі ЄЖООС файлу.
    /// </summary>
    public class MoneyReportWorkflow
    {
        private const string PremCaption = "Рапорт на премію";
        private const string AddRewardCaption = "Рапорт на ГЗ";

        private static readonly string ShiftEnter = ((char)11).ToString();
        private static readonly string[] UpdateStates = { "-", "\\", "|", "/" };

        private readonly Application excelApp;
        private int updateStateIndex;

        public MoneyReportWorkflow(Application excelApp)
        {
            this.excelApp = excelApp;
        }

        /// <summary>
        /// Спільна для всіх рапортів перевірка: книга відкрита і є ЄЖООС, тобто містить аркуші
        /// «1. ШПО», «2. ООС» і «6. Табель». Якщо ні — показує повідомлення і повертає false.
        /// </summary>
        private static bool TryGetEjoosSheets(Workbook workbook, string caption, out EjoosSheets sheets)
        {
            sheets = new EjoosSheets();
            if (workbook != null)
            {
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    if (PositionsTable.IsPositionsTable(ws))
                        sheets.Positions = ws;
                    else if (OOSTable.IsOOSTable(ws))
                        sheets.Oos = ws;
                    else if (Tabel.IsTabel(ws))
                        sheets.Tabel = ws;
                }
            }

            if (sheets.Positions != null && sheets.Oos != null && sheets.Tabel != null)
            {
                return true;
            }

            UserDialogs.ShowMessage(
                "Завантажте таблицю ЄЖООС, для якої треба створити рапорт " +
                $"(з аркушами «{PositionsTable.SheetName}», «{OOSTable.SheetName}», «{Tabel.SheetName}»). " +
                "Наразі створення рапорта неможливе.",
                caption, MessageBoxIcon.Error);
            return false;
        }

        /// <summary>
        /// Аркуші ЄЖООС книги.
        /// </summary>
        private class EjoosSheets
        {
            public Worksheet Positions { get; set; }
            public Worksheet Oos { get; set; }
            public Worksheet Tabel { get; set; }
        }

        #region Рапорт на премію

        /// <summary>
        /// Генерує рапорт на премію у Word документі (DocX) на основі таблиці ООС у ЄЖООС книзі.
        /// </summary>
        public void CreatePremReport(Workbook workbook)
        {
            if (!TryGetEjoosSheets(workbook, PremCaption, out var sheets))
            {
                return;
            }

            var oosTable = new OOSTable(sheets.Oos);

            if (!UserDialogs.TryChooseFolder("Оберіть папку для збереження рапорту на премію",
                    out var selectedFolder))
            {
                return;
            }

            var reportPath = Path.Combine(selectedFolder, "рапорт_на_премію.docx");
            WritePremDocument(oosTable, reportPath);

            Process.Start(new ProcessStartInfo(reportPath) { UseShellExecute = true });
        }

        private static void WritePremDocument(OOSTable oosTable, string reportPath)
        {
            using (var document = DocX.Create(reportPath))
            {
                var defaultFont = new XDoc.Font("Times New Roman");
                document.SetDefaultFont(defaultFont, 14);
                document.InsertParagraph("рапорт на премію");

                //Create table
                var table = document.AddTable(oosTable.Items.Count + 1, 7);

                // Create autonumbered list for first column
                var numList = document.AddList(null, 0, XDoc.ListItemType.Numbered, 1);
                for (var i = 0; i < oosTable.Items.Count; i++)
                {
                    document.AddListItem(numList, "");
                }

                //Add header
                table.Rows[0].Cells[0].Paragraphs[0].Append("№ п/п").FontSize(10);
                table.Rows[0].Cells[1].Paragraphs[0].Append("Військове звання").FontSize(10);
                table.Rows[0].Cells[2].Paragraphs[0].Append("Прізвище, ім’я,\nпо батькові").FontSize(10);
                table.Rows[0].Cells[3].Paragraphs[0].Append("Індекс посади").FontSize(10);
                table.Rows[0].Cells[4].Paragraphs[0].Append("РНОКПП").FontSize(10);
                table.Rows[0].Cells[5].Paragraphs[0].Append("Розмір премії").FontSize(10);
                table.Rows[0].Cells[6].Paragraphs[0].Append("Стягнення\n(ким, коли і за що накладено)")
                    .FontSize(10);

                //Add data
                for (var i = 0; i < oosTable.Items.Count; i++)
                {
                    var item = oosTable.Items[i];
                    var numberCell = table.Rows[i + 1].Cells[0];
                    numberCell.RemoveParagraphAt(0);
                    var listItem = numList.Items[i];
                    listItem.FontSize(10);
                    listItem.Append("").FontSize(10);
                    numberCell.InsertParagraph(listItem);
                    table.Rows[i + 1].Cells[1].Paragraphs[0].Append(item.Rank.RankName).FontSize(10);
                    table.Rows[i + 1].Cells[2].Paragraphs[0].Append(item.Person.ToString()).FontSize(10);
                    table.Rows[i + 1].Cells[3].Paragraphs[0].Append(item.Position.First().ToString() ?? "")
                        .FontSize(10);
                    table.Rows[i + 1].Cells[4].Paragraphs[0].Append(item.RNOKPP).FontSize(10);
                    table.Rows[i + 1].Cells[5].Paragraphs[0].Append("100%").FontSize(10);
                    table.Rows[i + 1].Cells[6].Paragraphs[0].Append("Відсутнє").FontSize(10);
                }

                document.InsertTable(table);

                //Save document
                document.Save();
            }
        }

        #endregion

        #region Рапорт на додаткову винагороду (100/30)

        /// <summary>
        /// Генерує у новому документі Word таблиці рапорту на додаткову винагороду (100k / 30k / 10k)
        /// на основі аркушів «ШПО» і «Табель» ЄЖООС книги.
        /// </summary>
        public void CreateAddRewardReport(Workbook workbook)
        {
            try
            {
                if (!TryGetEjoosSheets(workbook, AddRewardCaption, out var sheets))
                {
                    return;
                }

                var positions = PositionsTable.ReadPositionsTable(sheets.Positions);

                if (!TryReadTabel(sheets.Tabel, out var tabelTable))
                {
                    return;
                }

                var tabelPositions = MatchPositions(tabelTable, positions);
                var reportDate = ReadReportDate(sheets.Tabel);

                WriteAddRewardDocument(tabelPositions, reportDate);

                UserDialogs.ShowMessage(
                    "Успішно створено таблиці грошового рапорта в окремий документ Word. Додайте текст рапорта та збережіть документ.",
                    AddRewardCaption, MessageBoxIcon.Information);
            }
            finally
            {
                excelApp.StatusBar = false;
            }
        }

        private static bool TryReadTabel(Worksheet tabelWorksheet, out List<TabelEntry> tabelTable)
        {
            tabelTable = Tabel.ReadTabel(tabelWorksheet);
            if (tabelTable == null || tabelTable.Count == 0)
            {
                UserDialogs.ShowMessage("Нема данних в таблиці Табель. Створення рапорта на ГЗ неможливе.",
                    AddRewardCaption, MessageBoxIcon.Information);
                return false;
            }

            return true;
        }

        /// <summary>
        /// Зіставляє записи Табеля з посадами ШПО за ПІБ. Особи без посади попадають у рапорт, але з попередженням.
        /// </summary>
        private static List<TabelPositionPair> MatchPositions(List<TabelEntry> tabelTable,
            List<PositionEntry> positions)
        {
            var tabelPositions = new List<TabelPositionPair>();
            foreach (var entry in tabelTable)
            {
                var position = positions.FirstOrDefault(p => string.Equals((p.PersonName ?? "").Trim(),
                    (entry.FullName ?? "").Trim(), StringComparison.OrdinalIgnoreCase));
                if (position == null)
                {
                    UserDialogs.ShowMessage(
                        $"Не знайдена посада {entry.FullName} в таблиці ШПО. Перевірте правильність написання ПІБ та наявність цієї особи в ШПО. " +
                        $"Вона буде включена в рапорт, але її посада не буде вказана. {entry.FullName}",
                        AddRewardCaption, MessageBoxIcon.Warning);
                }

                tabelPositions.Add(new TabelPositionPair { Tabel = entry, Position = position });
            }

            return tabelPositions;
        }

        /// <summary>
        /// Дата рапорту з клітинки A3 аркуша Табель; якщо не розпізнано — поточна дата.
        /// </summary>
        private static DateTime ReadReportDate(Worksheet tabelWorksheet)
        {
            var reportDateRaw = tabelWorksheet.Cells[3, 1].Value;
            if (reportDateRaw is DateTime dt)
            {
                return dt;
            }

            if (reportDateRaw is double d)
            {
                return DateTime.FromOADate(d);
            }

            if (DateTime.TryParse(reportDateRaw?.ToString(), out DateTime result))
            {
                return result;
            }

            UserDialogs.ShowMessage(
                "Не вдалося розпізнати дату в таблиці Табель (клітинка A3). Очікується дата в числовому форматі день.місяць.рік. Наразі буде використана поточна дата.",
                AddRewardCaption, MessageBoxIcon.Information);
            return DateTime.Now;
        }

        private void WriteAddRewardDocument(IReadOnlyList<TabelPositionPair> tabelPositions, DateTime reportDate)
        {
            Word.Application wordApp = null;
            Word.Document doc = null;
            try
            {
                wordApp = new Word.Application();
                wordApp.Visible = true;
                doc = wordApp.Documents.Add();

                doc.Content.InsertAfter("100k");
                CreateMoneyReportForState(doc, tabelPositions, reportDate, "++");
                doc.Content.InsertParagraphAfter();
                doc.Content.InsertAfter("30k");
                CreateMoneyReportForState(doc, tabelPositions, reportDate, "+");
                doc.Content.InsertParagraphAfter();
                doc.Content.InsertAfter("10k");
                CreateMoneyReportForState(doc, tabelPositions, reportDate, "-");
            }
            finally
            {
                // do not quit Word so user can see the document; release COM objects if created
                if (doc != null) Marshal.ReleaseComObject(doc);
                if (wordApp != null) Marshal.ReleaseComObject(wordApp);
            }
        }

        private void CreateMoneyReportForState(Word.Document reportDoc, IReadOnlyList<TabelPositionPair> tabelPositions,
            DateTime reportDate, string stateString)
        {
            var warnings = new List<string>();

            // insert table at end of document
            var range = reportDoc.Content;
            range.Collapse(0); // wdCollapseEnd
            var table = reportDoc.Tables.Add(range, 1, 7);
            table.AllowAutoFit = false;
            table.Borders.Enable = 1;

            // Set table font to Times New Roman, size 11
            table.Range.Font.Name = "Times New Roman";
            table.Range.Font.Size = 10;

            // headers
            table.Cell(1, 1).Range.Text = "№ п/п";
            table.Cell(1, 2).Range.Text = "Посада";
            table.Cell(1, 3).Range.Text = "Звання";
            table.Cell(1, 4).Range.Text = "ПІБ";
            table.Cell(1, 5).Range.Text = "Початок виконання завдання";
            table.Cell(1, 6).Range.Text = "Завершення періоду виконання";
            table.Cell(1, 7).Range.Text = "Примітки";

            // fill rows: for each SubUnit group, write a merged SubUnit row, then all data rows in that group
            var lastCommonSubUnit = "";
            var subheaderRows = new List<Word.Row>(); // collect subheader rows for merging after population

            foreach (var tabelPosition in tabelPositions)
            {
                if (tabelPosition.Tabel.States.All(s => s.State != stateString))
                    continue;

                // Make subunit subheader
                if (tabelPosition.Position?.SubUnit != null && lastCommonSubUnit != tabelPosition.Position.SubUnit)
                {
                    lastCommonSubUnit = tabelPosition.Position.SubUnit;
                    var subunitRow = table.Rows.Add();
                    subheaderRows.Add(subunitRow); // collect for later merging
                    subunitRow.Cells[1].Range.Text = lastCommonSubUnit;
                    subunitRow.Cells[1].Range.ParagraphFormat.Alignment =
                        Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    subunitRow.Cells[1].VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    subunitRow.Range.Font.Color = Word.WdColor.wdColorAutomatic;
                }

                //Add row for entry
                var entryRow = table.Rows.Add();
                entryRow.Cells[2].Range.Text = tabelPosition.Position?.Name ?? "Посада не знайдена";
                entryRow.Cells[3].Range.Text = tabelPosition.Tabel.Rank;
                entryRow.Cells[4].Range.Text = tabelPosition.Tabel.FullName;
                var sequences = tabelPosition.Tabel.States.Where(s => s.State == stateString).ToArray();
                entryRow.Cells[5].Range.Text =
                    string.Join(ShiftEnter, sequences.Select(s => GetDateTimeString(s.StartDay, reportDate)));
                entryRow.Cells[6].Range.Text =
                    string.Join(ShiftEnter, sequences.Select(s => GetDateTimeString(s.EndDay, reportDate)));

                // fill cell 7 with various additional state types in chronological order: вдр, лік, від, ...
                var notes = new List<string>();
                foreach (var state in tabelPosition.Tabel.States)
                {
                    if (state.State != "-" && state.State != "+" && state.State != "++" &&
                        !string.IsNullOrEmpty(state.State))
                    {
                        string label;
                        if (state.State == "вдр")
                            label = "Відрядження";
                        else if (state.State == "лік")
                            label = "Лікування";
                        else if (state.State == "від")
                            label = "Відпустка";
                        else if (state.State == "ВП")
                            label = "Відпустка по пораненню";
                        else if (state.State == "влк")
                            label = "ВЛК";
                        else
                            label = state.State;

                        notes.Add(
                            $"{label} {GetDateTimeString(state.StartDay, reportDate)}-{GetDateTimeString(state.EndDay, reportDate)}");
                    }
                }

                if (notes.Any())
                {
                    entryRow.Cells[7].Range.Text = string.Join(ShiftEnter, notes);
                }

                //Some validations
                if (tabelPosition.Position == null)
                {
                    warnings.Add($"Для {tabelPosition.Tabel.FullName} не знайдена відповідна посада в ШПО.");
                    entryRow.Range.Font.Color = Word.WdColor.wdColorRed;
                }
                else if (tabelPosition.Tabel.ContainsEmptyRanges)
                {
                    warnings.Add($"Для {tabelPosition.Tabel.FullName} в Табелі існують незаповнені значення.");
                    entryRow.Range.Font.Color = Word.WdColor.wdColorOrange;
                }
                else
                {
                    entryRow.Range.Font.Color = Word.WdColor.wdColorAutomatic;
                }

                UpdateStatus("Створення рапорта на ГЗ...");
            }

            // Add autonumeration to 1st column but skip merged subheaders lines
            // Отримуем шаблон нумерації
            var template = reportDoc.Application.ListGalleries[Word.WdListGalleryType.wdNumberGallery]
                .ListTemplates[1];
            template.ListLevels[1].NumberFormat = "%1";
            template.ListLevels[1].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingNone;
            template.ListLevels[1].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            template.ListLevels[1].NumberPosition = 0;
            template.ListLevels[1].TextPosition = 0;

            var isFirstItem = true;
            for (var rowIndex = 2; rowIndex <= table.Rows.Count; rowIndex++)
            {
                var currentRow = table.Rows[rowIndex];

                // Skip if this is a subheader row
                if (subheaderRows.Any(r => r.Index == currentRow.Index))
                    continue;

                var firstCell = currentRow.Cells[1];

                // For normal data rows, add the number to first column
                firstCell.Range.ListFormat.ApplyListTemplateWithLevel(
                    ListTemplate: template,
                    ContinuePreviousList: !isFirstItem,
                    ApplyTo: Word.WdListApplyTo.wdListApplyToWholeList
                );
                isFirstItem = false;

                // Налаштування форматування першої колонки
                firstCell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                firstCell.Range.ParagraphFormat.LeftIndent = 0;
                firstCell.Range.ParagraphFormat.RightIndent = 0;
                firstCell.Range.ParagraphFormat.FirstLineIndent = 0;
            }

            // Merge cells in subheader rows after table population is complete
            foreach (var subheaderRow in subheaderRows)
            {
                subheaderRow.Cells[1].Merge(subheaderRow.Cells[7]);
                // ensure the merged cell keeps centered alignment
                subheaderRow.Cells[1].Range.ParagraphFormat.Alignment =
                    Word.WdParagraphAlignment.wdAlignParagraphCenter;
                subheaderRow.Cells[1].VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            }

            //Add warnings at the end of the document
            if (warnings.Any())
            {
                var warningRange = reportDoc.Content;
                warningRange.Collapse(0); // wdCollapseEnd
                warningRange.InsertParagraphAfter();
                warningRange.Collapse(0); // wdCollapseEnd
                warningRange.Text = "Попередження:\n" + string.Join("\n", warnings);
            }

            table.AllowAutoFit = true;
            table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent);
            table.AllowAutoFit = false;
        }

        private static string GetDateTimeString(int dayNumber, DateTime reportDate)
        {
            if (dayNumber < 1 || dayNumber > DateTime.DaysInMonth(reportDate.Year, reportDate.Month))
                throw new ArgumentOutOfRangeException(nameof(dayNumber));

            var date = new DateTime(reportDate.Year, reportDate.Month, dayNumber);
            return date.ToString("dd.MM.yyyy");
        }

        private void UpdateStatus(string operationName)
        {
            excelApp.StatusBar = $"{operationName} {UpdateStates[updateStateIndex++ % UpdateStates.Length]}";
        }

        private class TabelPositionPair
        {
            public TabelEntry Tabel { get; set; }
            public PositionEntry Position { get; set; }
        }

        #endregion
    }
}