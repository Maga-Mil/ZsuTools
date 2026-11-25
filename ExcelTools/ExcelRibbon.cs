using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Xceed.Document.NET;
using Xceed.Words.NET;
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
#if DEBUG
            const bool isDebug = true;
            const string tabId = "tabZsuTools_DEBUG";
            const string tabLabel = "ЗСУ debug";
#else
            const bool isDebug = false;
            const string tabId = "tabZsuTools";
            const string tabLabel = "ЗСУ";
#endif

            return new CustomUIBuilder()
                .AddTab(tabId, tabLabel)
                .AddGroup("groupMain", "ЄЖООС")
                .AddButton("btnAddReward", "Рапорт на додаткову винагороду", "OnAddRewardClicked",
                    "InternationalCurrency")
                .AddButton("btnPremReward", "Рапорт на премію", "OnPremRewardClicked", "InternationalCurrency")
                .AddGroup("groupStroyova", "Стройова")
                .AddButton("btnGetTabelFromStroyovaBatch", "Отримати табель зі Стройової Записки (пакетно)",
                    "OnGetTabelFromStroyovaBatchClicked", "AccessListEvents")
                .AddButtonIf(isDebug, "btnGetTabelFromStroyova", "Отримати табель зі Стройової Записки",
                    "OnGetTabelFromStroyovaClicked", "AccessListEvents")
                .AddGroup("tabelZaluchenosti", "Табель залученості")
                .AddButton("btnTabelZaluchenosti", "Оновити Табель Залученості", "OnGetTabelZaluchenostiClicked",
                    "AccessListEvents", supertip: "Заповнити Табель Залученості зі Стройової Записки")
                .AddButtonIf(isDebug, "btnTestTabelZaluchenosti", "Тест: порівняти COM і ClosedXML",
                    "OnTestTabelZaluchenostiClicked", "AccessListEvents")
                .AddButtonIf(isDebug, "btnTestSaveTabelZaluchenosti", "Тест: збереження Job/JobDetails",
                    "OnTestSaveTabelZaluchenostiClicked", "AccessListEvents")
                .Build();
        }

        // Called when Excel loads the ribbon
        public void OnLoad(IRibbonUI ribbonUi)
        {
            _ribbon = ribbonUi;
        }

        /// <summary>
        /// Сгенерувати рапорт на премію у Word документі на основі таблиці ООС у активному ЄЖООС файлі.
        /// </summary>
        /// <param name="control"></param>
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
                                    document.AddListItem(numList, "");
                                }

                                //Add header
                                table.Rows[0].Cells[0].Paragraphs[0].Append("№ п/п").FontSize(10);
                                table.Rows[0].Cells[1].Paragraphs[0].Append("Військове звання").FontSize(10);
                                table.Rows[0].Cells[2].Paragraphs[0].Append("Прізвище, ім’я,\nпо батькові")
                                    .FontSize(10);
                                table.Rows[0].Cells[3].Paragraphs[0].Append("Індекс посади").FontSize(10);
                                table.Rows[0].Cells[4].Paragraphs[0].Append("РНОКПП").FontSize(10);
                                table.Rows[0].Cells[5].Paragraphs[0].Append("Розмір премії").FontSize(10);
                                table.Rows[0].Cells[6].Paragraphs[0].Append("Стягнення\n(ким, коли і за що накладено)")
                                    .FontSize(10);

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
                                    table.Rows[i + 1].Cells[2].Paragraphs[0].Append(item.Person.ToString())
                                        .FontSize(10);
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

            MessageBox.Show(ExcelWindow.Instance, $"Не можу знайти таблицю ООС.", "ZSUTools", MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }

        /// <summary>
        /// Сгенерувати рапорт на додаткову винагороду у Word документі на основі табеля у активному ЄЖООС файлі.
        /// </summary>
        /// <param name="control"></param>
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
                MessageBox.Show(ExcelWindow.Instance, $"Error catched, money report is discarded. {e}", "MoneyReport",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                excelApp.StatusBar = false;
                //excelApp.DisplayStatusBar = false;
            }
        }

        /// <summary>
        /// Отримати табель зі Стройової Записки у активному ЄЖООС файлі. Поки що це здебільшого тестовий код.
        /// TODO Зробити заповнення Табеля ЄЖООС на основі Стройової Записки, а не просто вивід статистики у MessageBox.
        /// </summary>
        /// <param name="control"></param>
        public void OnGetTabelFromStroyovaClicked(IRibbonControl control)
        {
            Application excelApp = (Application)ExcelDnaUtil.Application;
            try
            {
                //var stroyovaCOM = new СтройоваЗаписка(excelApp.ActiveWorkbook);

                // Читаємо Стройову Записку з файлу через ClosedXML (без Excel Interop)
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "Файли Excel (*.xlsx)|*.xlsx|Усі файли (*.*)|*.*";
                    openFileDialog.Title = "Оберіть файл Стройової Записки";

                    if (openFileDialog.ShowDialog() != DialogResult.OK)
                    {
                        return;
                    }

                    var stroyova = new StroyovaZapyska(openFileDialog.FileName);

                    var byLocation = stroyova.Items
                        .GroupBy(t => t.State.Location)
                        .Select(g => $"{g.Key}: {g.Count()}");
                    var sample = stroyova.Items
                        .Select(t => $"{t.Person.FullName} | {t.State.Location} | {t.Color.Name}");

                    MessageBox.Show(ExcelWindow.Instance,
                        $"Прочитано записів: {stroyova.Items.Count}\n" +
                        $"{string.Join("\n", byLocation)}\n\n" +
                        $"Перші записи:\n{string.Join("\n", sample)}",
                        "Стройова записка (ClosedXML)", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(ExcelWindow.Instance,
                    $"Виникла помилка, створення табеля зі Стройової записки перервано. {e}",
                    "Стройова записка", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                excelApp.StatusBar = false;
                //excelApp.DisplayStatusBar = false;
            }
        }

        /// <summary>
        /// Отримати табель зі Стройової Записки у активному ЄЖООС файлі, обробляючи всі файли місяця у папці.
        /// Це не Табель ЄЖООС, а просто табличка схожа на Табель
        /// </summary>
        /// <param name="control"></param>
        public void OnGetTabelFromStroyovaBatchClicked(IRibbonControl control)
        {
            Application excelApp = (Application)ExcelDnaUtil.Application;
            try
            {
                var workBook = excelApp.ActiveWorkbook;
                //excelApp.DisplayStatusBar = true;

                var batchResult = BatchProcessor.ProcessMonthlyFiles();
                ResultTableGenerator.GenerateSummaryTable(batchResult);
                MessageBox.Show(ExcelWindow.Instance,
                    $"Табель зі Стройової Записки сформовано успішно, оброблено {batchResult.Count} файлів.",
                    "Стройова записка", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception e)
            {
                MessageBox.Show(ExcelWindow.Instance,
                    $"Виникла помилка, створення табеля зі Стройової записки перервано. {e}",
                    "Стройова записка", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                excelApp.StatusBar = false;
                //excelApp.DisplayStatusBar = false;
            }
        }

        /// <summary>
        /// Ручний тест: порівнює читання Табеля Залученості через COM (активна книга) і ClosedXML (файл з діалогу).
        /// </summary>
        public void OnTestTabelZaluchenostiClicked(IRibbonControl control)
        {
            Tests.TabelZaluchenostiTests.CompareComAndClosedXml();
        }

        /// <summary>
        /// Ручний тест: зміна Job/JobDetails, збереження через COM, перечитування і відновлення початкових значень.
        /// </summary>
        public void OnTestSaveTabelZaluchenostiClicked(IRibbonControl control)
        {
            Tests.TabelZaluchenostiTests.SaveRoundTrip();
        }

        /// <summary>
        /// Заповнює Залученість (Job) у Табелі Залученості активної книги на основі обраної Стройової Записки:
        /// БЧ → BATTALION_DEFENSE_SECTOR, РЗ → COMBAT_SUPPORT. Відсутніх не змінює.
        /// </summary>
        public void OnGetTabelZaluchenostiClicked(IRibbonControl control)
        {
            Application excelApp = (Application)ExcelDnaUtil.Application;
            try
            {
                var workBook = excelApp.ActiveWorkbook;
                if (workBook == null)
                {
                    MessageBox.Show(ExcelWindow.Instance, "Не завантажено Табель Залученості.", "Табель залученості");
                    return;
                }

                var ws = ExcelUtils.FindWorksheetByNameContains(workBook, "Табелювання");
                if (ws == null)
                {
                    MessageBox.Show(ExcelWindow.Instance,
                        "Не знайдено таблицю 'Табелювання', це не схоже на Табель Залученості", "Табель залученості");
                    return;
                }

                var tz = new TabelZaluchenosti(workBook);

                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "Файли Excel (*.xlsx)|*.xlsx|Усі файли (*.*)|*.*";
                    openFileDialog.Title = "Оберіть Стройову Записку, щоб скопіювати дані до Табеля Залученості";

                    if (openFileDialog.ShowDialog() != DialogResult.OK)
                    {
                        return;
                    }

                    string selectedFilePath = openFileDialog.FileName;

                    StroyovaZapyska sz;
                    var szWb = excelApp.Workbooks.Open(selectedFilePath, ReadOnly: true);
                    try
                    {
                        sz = new StroyovaZapyska(szWb);
                    }
                    finally
                    {
                        szWb.Close(false);
                    }

                    // Табель міг перестати бути активною книгою, поки відкривалася Стройова
                    workBook.Activate();

                    // Some sanity checks
                    if (tz.Items.Count != sz.Items.Count)
                    {
                        MessageBox.Show(ExcelWindow.Instance,
                            $"Кількість записів у Табелі Залученості ({tz.Items.Count}) не збігається з кількістю записів у Стройовій Записці ({sz.Items.Count}). Операця буде зупинена.",
                            "Табель залученості");
                        return;
                    }

                    var mismatchedRecords = new List<string>();
                    foreach (var tzItem in tz.Items)
                    {
                        var person = sz.Items.SingleOrDefault(r =>
                            r.Person.FullName.Equals(tzItem.PersonFullName,
                                StringComparison.InvariantCultureIgnoreCase));
                        if (person == null)
                        {
                            mismatchedRecords.Add(tzItem.PersonFullName);
                        }
                    }

                    if (mismatchedRecords.Count > 0)
                    {
                        MessageBox.Show(ExcelWindow.Instance,
                            $"У Табелі Залученості присутні записи, які відсутні у Стройовій Записці:\n\r {String.Join(", ", mismatchedRecords)}. \n\rТреба виправити ПІБ у Стройовій Записці відповідно до Табеля Залученості. Операця буде зупинена.",
                            "Табель залученості");
                        return;
                    }

                    // Для кожної особи зі Стройової знаходимо відповідну особу в Табелі (за ПІБ)
                    var tzByName = tz.Items.ToLookup(i => i.PersonFullName, StringComparer.InvariantCultureIgnoreCase);
                    var notFound = new List<string>();
                    var ambiguous = new List<string>();
                    var unknownLocations = new List<string>();
                    var missingJobValues = new List<string>();
                    var assignments = new List<KeyValuePair<TabelZaluchenosti.Item, TabelZaluchenosti.JobValue>>();

                    foreach (var szItem in sz.Items)
                    {
                        var job = LocationToJob(szItem.State, tz.Jobs, out string jobKey);
                        if (job == null)
                        {
                            unknownLocations.Add($"{szItem.Person.FullName} — «{szItem.State.LocationName}»");
                            continue;
                        }

                        if (ReferenceEquals(job, TabelZaluchenosti.JobValue.Unknown))
                        {
                            // Значення Залученості відсутнє у довіднику InvVal Табеля
                            if (!missingJobValues.Contains(jobKey))
                            {
                                missingJobValues.Add(jobKey);
                            }

                            continue;
                        }

                        var matches = tzByName[szItem.Person.FullName].ToList();
                        if (matches.Count == 0)
                        {
                            notFound.Add(szItem.Person.FullName);
                        }
                        else if (matches.Count > 1)
                        {
                            ambiguous.Add(szItem.Person.FullName);
                        }
                        else
                        {
                            assignments.Add(new KeyValuePair<TabelZaluchenosti.Item, TabelZaluchenosti.JobValue>(
                                matches[0], job));
                        }
                    }

                    if (notFound.Count > 0 || ambiguous.Count > 0 || unknownLocations.Count > 0 ||
                        missingJobValues.Count > 0)
                    {
                        var message = new StringBuilder();
                        if (unknownLocations.Count > 0)
                        {
                            message.AppendLine(
                                $"Невідоме місцезнаходження у Стройовій Записці (виправте Стройову Записку):\n{string.Join("\n", unknownLocations)}\n");
                        }

                        if (missingJobValues.Count > 0)
                        {
                            message.AppendLine(
                                $"У довіднику Залученості (InvVal) Табеля відсутні значення:\n{string.Join(", ", missingJobValues)}\n");
                        }

                        if (notFound.Count > 0)
                        {
                            message.AppendLine(
                                $"Особи зі Стройової Записки, яких не знайдено у Табелі Залученості:\n{string.Join(", ", notFound)}\n");
                        }

                        if (ambiguous.Count > 0)
                        {
                            message.AppendLine(
                                $"ПІБ, які зустрічаються у Табелі Залученості кілька разів:\n{string.Join(", ", ambiguous)}\n");
                        }

                        message.Append("Операція буде зупинена, Табель Залученості не змінено.");
                        MessageBox.Show(ExcelWindow.Instance, message.ToString(), "Табель залученості");
                        return;
                    }

                    foreach (var assignment in assignments)
                    {
                        assignment.Key.Job = assignment.Value;
                    }

                    int written = tz.Save();

                    var result = new StringBuilder();
                    result.AppendLine($"Залученість заповнено зі Стройової Записки на {sz.Date:dd.MM.yyyy}.");
                    result.AppendLine($"Оброблено осіб: {assignments.Count}, змінено комірок: {written}.");

                    var warnings = tz.GetPresenceJobWarnings();
                    if (warnings.Count > 0)
                    {
                        result.AppendLine();
                        result.AppendLine(
                            $"Увага! Наявність не узгоджується із Залученістю ({warnings.Count}), перевірте:");
                        foreach (var warning in warnings)
                        {
                            result.AppendLine(warning);
                        }
                    }

                    MessageBox.Show(ExcelWindow.Instance, result.ToString(), "Табель залученості",
                        MessageBoxButtons.OK,
                        warnings.Count > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(ExcelWindow.Instance, $"Помилка при завантаженні табеля залученості: {e}",
                    "Табель залученості");
            }
        }

        /// <summary>
        /// Відповідність місцезнаходження зі Стройової Записки значенню Залученості Табеля.
        /// Повертає null, якщо місцезнаходження невідоме (треба виправити Стройову Записку).
        /// У jobKey — код значення Залученості (для повідомлення, якщо його немає у довіднику InvVal).
        /// </summary>
        private static TabelZaluchenosti.JobValue LocationToJob(StroyovaZapyska.Place place,
            TabelZaluchenosti.JobValue.Registry jobs, out string jobKey)
        {
            switch (place.Location)
            {
                case StroyovaZapyska.LocationType.БЧ:
                    jobKey = nameof(jobs.BATTALION_DEFENSE_SECTOR);
                    return jobs.BATTALION_DEFENSE_SECTOR;
                case StroyovaZapyska.LocationType.РЗ:
                    jobKey = nameof(jobs.COMBAT_SUPPORT);
                    return jobs.COMBAT_SUPPORT;
            }

            switch ((place.LocationName ?? string.Empty).Trim())
            {
                case "Відрядження":
                    jobKey = nameof(jobs.TRIP);
                    return jobs.TRIP;
                case "Відпустка щорічна":
                case "Відпустка ЗСО":
                case "Відпустка ЗСЗ":
                    jobKey = nameof(jobs.VACATION);
                    return jobs.VACATION;
                case "СЗЧ":
                    jobKey = nameof(jobs.AWOL);
                    return jobs.AWOL;
                case "Лікування (до 10 діб)":
                    jobKey = nameof(jobs.SHORT_TERM_TREATMENT);
                    return jobs.SHORT_TERM_TREATMENT;
                case "Лікування довготривале":
                    jobKey = nameof(jobs.LONG_TERM_TREATMENT);
                    return jobs.LONG_TERM_TREATMENT;
                default:
                    jobKey = null;
                    return null;
            }
        }


        private void UpdateStatus(Application excelApp, string operationName)
        {
            excelApp.StatusBar = $"{operationName} {UpdateStates[_updateStateIndex++ % UpdateStates.Length]}";
        }

        private static readonly string[] UpdateStates = new[] { "-", "\\", "|", "/" };
        private int _updateStateIndex = 0;
    }

    /// <summary>
    /// Wraps Excel's main HWND handle into an IWin32Window implementation
    /// so WinForms dialogs and MessageBoxes bind correctly as modal windows to Excel.
    /// </summary>
    public class ExcelWindow : IWin32Window
    {
        public IntPtr Handle => ExcelDnaUtil.WindowHandle;

        /// <summary>
        /// Gets a reusable singleton instance of the Excel main window wrapper.
        /// </summary>
        public static ExcelWindow Instance { get; } = new ExcelWindow();
    }
}