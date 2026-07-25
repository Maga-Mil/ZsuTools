using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using ZsuTools.Tables;
using Word = Microsoft.Office.Interop.Word;

namespace ZsuTools
{
    public static class MoneyReportGenerator
    {
       public static void CreateMoney_100_30_Report( Workbook activeWb, Action<string> onUpdate = null )
        {
            if ( activeWb == null )
            {
                MessageBox.Show( "Завантажте таблицю ЄЖООС для якої треба створити рапорт на грошове забезпечення. Наразі створення рапорта неможливе.", 
                        "MoneyReport", MessageBoxButtons.OK, MessageBoxIcon.Error );
                return;
            };

            // Read positions (ШПО)
            var positionsSheet = ExcelUtils.FindWorksheetByNameContains( activeWb, "ШПО" );
            if ( positionsSheet == null )
            {
                MessageBox.Show( "Не найдена таблиця ШПО. Створення рапорта на ГЗ неможливе.", 
                        "MoneyReport", MessageBoxButtons.OK, MessageBoxIcon.Error );
                return;
            }
            var positions      = PositionsTable.ReadPositionsTable( positionsSheet );

            // Read table (Табель)
            var tabelWorksheet = ExcelUtils.FindWorksheetByNameContains( activeWb, "табель" );
            if ( tabelWorksheet == null )
            {
                MessageBox.Show( "Не найдена таблиця Табель. Створення рапорта на ГЗ неможливе.", 
                        "MoneyReport", MessageBoxButtons.OK, MessageBoxIcon.Error );
                return;
            }

            var tabelTable = Tabel.ReadTabel( tabelWorksheet );
            if ( tabelTable == null || tabelTable.Count == 0 )
            {
                MessageBox.Show( "Нема данних в таблиці Табель. Створення рапорта на ГЗ неможливе.", 
                        "MoneyReport", MessageBoxButtons.OK, MessageBoxIcon.Information );
                return;
            }

            //Make TabelPositionPair list for easier matching
            var tabelPositions = new List<TabelPositionPair>();
            foreach ( var entry in tabelTable )
            {
                var position = positions.FirstOrDefault( p => string.Equals( (p.PersonName ?? "").Trim(), (entry.FullName ?? "").Trim(),
                        StringComparison.OrdinalIgnoreCase ) );
                if( position == null )
                {
                    MessageBox.Show(
                            $"Не знайдена посада {entry.FullName} в таблиці ШПО. Перевірте правильність написання ПІБ та наявність цієї особи в ШПО. " +
                            $"Вона буде включена в рапорт, але її посада не буде вказана. {entry.FullName}",
                            "MoneyReport", MessageBoxButtons.OK, MessageBoxIcon.Warning );
                }

                var tabelPositionPair = new TabelPositionPair
                                        {
                                                Tabel    = entry,
                                                Position = position
                                        };
                tabelPositions.Add( tabelPositionPair );
            }

            //Get current report date from A3 cell of the Tabel worksheet
            DateTime reportDate    = DateTime.Now;
            var      reportDateRaw = tabelWorksheet.Cells[ 3, 1 ].Value;
            if( reportDateRaw is DateTime dt )
                reportDate = dt;
            else if ( reportDateRaw is double d )
                reportDate = DateTime.FromOADate( d );
            else if ( DateTime.TryParse( reportDateRaw?.ToString(), out DateTime result ) )
                reportDate = result;
            else
            {
                MessageBox.Show( "Не вдалося розпізнати дату в таблиці Табель (клітинка A3). Очікується дата в числовому форматі день.місяць.рік. Наразі буде використана поточна дата.", 
                        "MoneyReport", MessageBoxButtons.OK, MessageBoxIcon.Information );
            }

            Word.Application wordApp = null;
            Document         doc     = null;
            try
            {
                wordApp              = new Word.Application();
                wordApp.Visible      = true;
                doc                  = wordApp.Documents.Add();

                doc.Content.InsertAfter( "100k" );
                CreateMoneyReportForState( doc, tabelPositions, reportDate, "++", onUpdate );
                doc.Content.InsertParagraphAfter();
                doc.Content.InsertAfter( "30k" );
                CreateMoneyReportForState( doc, tabelPositions, reportDate, "+", onUpdate );

                MessageBox.Show( "Успішно створено таблиці грошового рапорта в окремий документ Word. Додайте текст рапорта та збережіть документ.",
                        "MoneyReport", MessageBoxButtons.OK, MessageBoxIcon.Information );
            }
            catch ( Exception ex )
            {
                MessageBox.Show( "Помилка створення рапорта на ГЗ: " + ex.Message, "MoneyReport", MessageBoxButtons.OK, MessageBoxIcon.Error );
                throw;
            }
            finally
            {
                // do not quit Word so user can see the document; release COM objects if created
                if ( doc     != null ) Marshal.ReleaseComObject( doc );
                if ( wordApp != null ) Marshal.ReleaseComObject( wordApp );
            }
        }

        private static void CreateMoneyReportForState( Document reportDoc, IReadOnlyList<TabelPositionPair> tabelPositions, DateTime reportDate,
                                                string   stateString, Action<string> onUpdate = null )
        {
            var warnings = new List<string>();

            // insert table at end of document
            var range = reportDoc.Content;
            range.Collapse( 0 ); // wdCollapseEnd
            var table = reportDoc.Tables.Add( range, 1, 7 );
            table.AllowAutoFit   = false;
            table.Borders.Enable = 1;

            // Set table font to Times New Roman, size 11
            table.Range.Font.Name = "Times New Roman";
            table.Range.Font.Size = 10;

            // headers
            table.Cell( 1, 1 ).Range.Text = "№ п/п";
            table.Cell( 1, 2 ).Range.Text = "Посада";
            table.Cell( 1, 3 ).Range.Text = "Звання";
            table.Cell( 1, 4 ).Range.Text = "ПІБ";
            table.Cell( 1, 5 ).Range.Text = "Початок виконання завдання";
            table.Cell( 1, 6 ).Range.Text = "Завершення періоду виконання";
            table.Cell( 1, 7 ).Range.Text = "Примітки";

            // fill rows: for each SubUnit group, write a merged SubUnit row, then all data rows in that group
            int itemIndex         = 1;
            var lastCommonSubUnit = "";
            var subheaderRows     = new List<Row>(); // collect subheader rows for merging after population

            foreach ( var tabelPosition in tabelPositions )
            {
                if( tabelPosition.Tabel.States.All( s => s.State != stateString ) )
                    continue;

                // Make subunit subheader
                if( tabelPosition.Position?.SubUnit != null && lastCommonSubUnit != tabelPosition.Position.SubUnit )
                {
                    lastCommonSubUnit = tabelPosition.Position.SubUnit;
                    var subunitRow = table.Rows.Add();
                    subheaderRows.Add( subunitRow ); // collect for later merging
                    subunitRow.Cells[ 1 ].Range.Text                      = lastCommonSubUnit;
                    subunitRow.Cells[ 1 ].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    subunitRow.Cells[ 1 ].VerticalAlignment               = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    subunitRow.Range.Font.Color                           = WdColor.wdColorAutomatic;
                }

                //Add row for entry
                var entryRow = table.Rows.Add();
                entryRow.Cells[ 1 ].Range.Text = itemIndex.ToString();
                entryRow.Cells[ 2 ].Range.Text = tabelPosition.Position?.Name ?? "Посада не знайдена";
                entryRow.Cells[ 3 ].Range.Text = tabelPosition.Tabel.Rank;
                entryRow.Cells[ 4 ].Range.Text = tabelPosition.Tabel.FullName;
                var sequences = tabelPosition.Tabel.States.Where( s => s.State == stateString ).ToArray();
                entryRow.Cells[ 5 ].Range.Text =
                        String.Join( ShiftEnter, sequences.Select( s => GetDateTimeString( s.StartDay, reportDate ) ) );
                entryRow.Cells[ 6 ].Range.Text =
                        String.Join( ShiftEnter, sequences.Select( s => GetDateTimeString( s.EndDay, reportDate ) ) );

                // fill cell 7 with various additional state types in chronological order: вдр, лік, від, ...
                var notes = new List<string>();
                foreach ( var state in tabelPosition.Tabel.States )
                {
                    if (state.State != "-" && state.State != "+" && state.State != "++" && !string.IsNullOrEmpty(state.State))
                    {
                        string label = null;
                        if ( state.State == "вдр" )
                            label = "Відрядження";
                        else if ( state.State == "лік" )
                            label = "Лікування";
                        else if ( state.State == "від" )
                            label = "Відпустка";
                        else if ( state.State == "ВП" )
                            label = "Відпустка по пораненню";
                        else if ( state.State == "влк" )
                            label = "ВЛК";
                        else 
                            label = state.State;

                        if ( label != null )
                        {
                            notes.Add( $"{label} {GetDateTimeString( state.StartDay, reportDate )}-{GetDateTimeString( state.EndDay, reportDate )}" );
                        }
                    }
                }

                if ( notes.Any() )
                {
                    entryRow.Cells[ 7 ].Range.Text = String.Join( ShiftEnter, notes );
                }

                //Some validations
                if( tabelPosition.Position == null )
                {
                    warnings.Add( $"Для {tabelPosition.Tabel.FullName} не знайдена відповідна посада в ШПО." );
                    entryRow.Range.Font.Color = WdColor.wdColorRed;
                }
                else
                {
                    entryRow.Range.Font.Color = WdColor.wdColorAutomatic;
                    if( tabelPosition.Tabel.ContainsEmptyRanges )
                    {
                        warnings.Add( $"Для {tabelPosition.Tabel.FullName} в Табелі існують незаповнені значення." );
                        entryRow.Range.Font.Color = WdColor.wdColorOrange;
                    }
                    else
                    {
                        entryRow.Range.Font.Color = WdColor.wdColorAutomatic;
                    }
                }

                itemIndex++;

                onUpdate?.Invoke("Generating money report...");
            }

            // Merge cells in subheader rows after table population is complete
            foreach ( var subheaderRow in subheaderRows )
            {
                subheaderRow.Cells[ 1 ].Merge( subheaderRow.Cells[ 7 ] );
                // ensure the merged cell keeps centered alignment
                subheaderRow.Cells[ 1 ].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                subheaderRow.Cells[ 1 ].VerticalAlignment               = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            }

            //Add warnings at the end of the document
            if ( warnings.Any() )            
            {
                var warningRange = reportDoc.Content;
                warningRange.Collapse( 0 ); // wdCollapseEnd
                warningRange.InsertParagraphAfter();
                warningRange.Collapse( 0 ); // wdCollapseEnd
                warningRange.Text = "Попередження:\n" + String.Join( "\n", warnings );
            }

            table.AllowAutoFit = true;
            table.AutoFitBehavior( WdAutoFitBehavior.wdAutoFitContent );
            table.AllowAutoFit = false; 
        }

        

        private static string GetDateTimeString( int dayNumber, DateTime reportDate )
        {
            if ( dayNumber < 1 || dayNumber > DateTime.DaysInMonth( reportDate.Year, reportDate.Month ) )
                throw new ArgumentOutOfRangeException( nameof(dayNumber) );

            DateTime date = new DateTime( reportDate.Year, reportDate.Month, dayNumber );
            return date.ToString( "dd.MM.yyyy" );
        }

        private static readonly string ShiftEnter = ((char)11).ToString();

        private class TabelPositionPair
        {
            public TabelEntry    Tabel    { get; set; }
            public PositionEntry Position { get; set; }
        }
    }
}