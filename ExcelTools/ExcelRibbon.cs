using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Xceed.Document.NET;
using Xceed.Words.NET;
using ZsuTools.GenerateTabel;
using ZsuTools.Tables;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace ZsuTools
{
    [ComVisible(true)]
    public class ZsuToolsRibbon : ExcelRibbon
    {
        private IRibbonUI _ribbon;

        // Return the ribbon XML for Excel to render.
        public override string GetCustomUI(string ribbonId)
        {
            var customUI = @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>"

#if DEBUG
                           + "<tab id='tabZsuTools_DEBUG' label='ЗСУ debug'>"
#else
            +"<tab id='tabZsuTools' label='ЗСУ'>"
#endif
                           +
                           @"<group id='groupMain' label='ЄЖООС'>"
                           + "<button id='btnAddReward' label='Рапорт на додаткову винагороду' size='large' imageMso='InternationalCurrency' showImage='true' onAction='OnAddRewardClicked' />"
//#if DEBUG
                           + "<button id='btnPremReward' label='Рапорт на премію' size='large' imageMso='InternationalCurrency' showImage='true' onAction='OnPremRewardClicked' />"
//#endif
                           + "</group>"
                           +
                           @"<group id='groupStroyova' label='Стройова'>
          <button id='btnGetTabelFromStroyovaBatch' label='Отримати табель зі Стройової Записки (пакетно)' size='large' imageMso='AccessListEvents' showImage='true' onAction='OnGetTabelFromStroyovaBatchClicked' />
"
// #if DEBUG
//                    + 
//         @"<button id='btnGetTabelFromStroyova' label='Отримати табель зі Стройової Записки' size='large' imageMso='AccessListEvents' showImage='true' onAction='OnGetTabelFromStroyovaClicked' />"
// #endif
                           +
                           @"</group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

            return customUI;
        }

        // Called when the ribbon is loaded by Excel
        public void OnLoad(IRibbonUI ribbonUi)
        {
            _ribbon = ribbonUi;
        }

        public void OnPremRewardClicked(IRibbonControl control)
        {
            Application excelApp = (Application)ExcelDnaUtil.Application;
            var workBook = excelApp.ActiveWorkbook;
            foreach (Microsoft.Office.Interop.Excel.Worksheet workBookSheet in workBook.Sheets)
            {
                if (OOSTable.IsOOSTable(workBookSheet))
                {
                    //Load ООС table
                    var oosTable = new OOSTable(workBookSheet);

                    //Open folder dialog
                    using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
                    {
                        // Optional configuration
                        folderDialog.Description = "Оберіть папку для збереження рапорту на премію";
                        folderDialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                        // Show the dialog and check if the user clicked OK
                        if (folderDialog.ShowDialog() == DialogResult.OK)
                        {
                            var selectedFolder = folderDialog.SelectedPath;
                            
                            //Create report base
                            var reportPath = Path.Combine(selectedFolder, "рапорт_на_премію.docx");
                            using (var document = DocX.Create(reportPath))
                            {
                                var defaultFont = new Font("Times New Roman");
                                document.SetDefaultFont(defaultFont, 14);
                                document.InsertParagraph("рапорт на премію");

                                //Create table
                                var table = document.AddTable(oosTable.Items.Count + 1, 7);
                                //table.Design = Xceed.Words.NET.TableDesign.LightShadingAccent1;

                                // Create autonumbered list for first column
                                var numList = document.AddList(null, 0, ListItemType.Numbered, 1);
                                for (int i = 0; i < oosTable.Items.Count; i++)
                                {
                                    document.AddListItem(numList,"");
                                }
                                
                                //Add header
                                table.Rows[0].Cells[0].Paragraphs[0].Append("№ п/п").FontSize(10);
                                table.Rows[0].Cells[1].Paragraphs[0].Append("Військове звання").FontSize(10);
                                table.Rows[0].Cells[2].Paragraphs[0].Append("Прізвище, ім’я,\nпо батькові").FontSize(10);
                                table.Rows[0].Cells[3].Paragraphs[0].Append("Індекс посади").FontSize(10);
                                table.Rows[0].Cells[4].Paragraphs[0].Append("РНОКПП").FontSize(10);
                                table.Rows[0].Cells[5].Paragraphs[0].Append("Розмір премії").FontSize(10);
                                table.Rows[0].Cells[6].Paragraphs[0].Append("Стягнення\n(ким, коли і за що накладено)").FontSize(10);

                                //Add data
                                for (int i = 0; i < oosTable.Items.Count; i++)
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
                                    table.Rows[i + 1].Cells[3].Paragraphs[0]
                                        .Append(item.Position.First().ToString() ?? "").FontSize(10);
                                    table.Rows[i + 1].Cells[4].Paragraphs[0].Append(item.RNOKPP).FontSize(10);
                                    table.Rows[i + 1].Cells[5].Paragraphs[0].Append("100%").FontSize(10);
                                    table.Rows[i + 1].Cells[6].Paragraphs[0].Append("Відсутнє").FontSize(10);
                                    
                                }

                                document.InsertTable(table);
                                

                                //Save document
                                document.Save();
                                
                                Process.Start(new ProcessStartInfo(reportPath) { UseShellExecute = true });
                            }
                        }
                    }


                    return;
                }
            }

            MessageBox.Show($"Не можу знайти таблицю ООС.", "ZSUTools", MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }

        public void OnAddRewardClicked(IRibbonControl control)
        {
            Application excelApp = (Application)ExcelDnaUtil.Application;
            try
            {
                var workBook = excelApp.ActiveWorkbook;
                //excelApp.DisplayStatusBar = true;

                void UpdateStatusFunc(String status)
                {
                    UpdateStatus(excelApp, status);
                }

                MoneyReportGenerator.CreateMoney_100_30_Report(workBook, UpdateStatusFunc);
            }
            catch (Exception e)
            {
                MessageBox.Show($"Error catched, money report is discarded. {e}", "MoneyReport", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                excelApp.StatusBar = false;
                //excelApp.DisplayStatusBar = false;
            }
        }

        public void OnGetTabelFromStroyovaClicked(IRibbonControl control)
        {
            Application excelApp = (Application)ExcelDnaUtil.Application;
            try
            {
                var workBook = excelApp.ActiveWorkbook;
                //excelApp.DisplayStatusBar = true;

                var stroyova = new СтройоваЗаписка(workBook);
            }
            catch (Exception e)
            {
                MessageBox.Show($"Виникла помилка, створення табеля зі Стройової записки перервано. {e}",
                    "Стройова записка", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                excelApp.StatusBar = false;
                //excelApp.DisplayStatusBar = false;
            }
        }

        public void OnGetTabelFromStroyovaBatchClicked(IRibbonControl control)
        {
            Application excelApp = (Application)ExcelDnaUtil.Application;
            try
            {
                var workBook = excelApp.ActiveWorkbook;
                //excelApp.DisplayStatusBar = true;

                var batchResult = BatchProcessor.ProcessMonthlyFiles();
                ResultTableGenerator.GenerateSummaryTable(batchResult);
                MessageBox.Show(
                    $"Табель зі Стройової Записки сформовано успішно, оброблено {batchResult.Count} файлів.",
                    "Стройова записка", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception e)
            {
                MessageBox.Show($"Виникла помилка, створення табеля зі Стройової записки перервано. {e}",
                    "Стройова записка", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                excelApp.StatusBar = false;
                //excelApp.DisplayStatusBar = false;
            }
        }


        private void UpdateStatus(Application excelApp, string operationName)
        {
            excelApp.StatusBar = $"{operationName} {UpdateStates[_updateStateIndex++ % UpdateStates.Length]}";
        }

        private static readonly string[] UpdateStates = new[] { "-", "\\", "|", "/" };
        private int _updateStateIndex = 0;
    }
}