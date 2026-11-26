using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ZsuTools.Tables;
using ZsuTools.UI;
using ZsuTools.Workflows;
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
        public void OnPremRewardClicked(IRibbonControl control)
        {
            Application excelApp = (Application)ExcelDnaUtil.Application;
            try
            {
                new MoneyReportWorkflow(excelApp).CreatePremReport(excelApp.ActiveWorkbook);
            }
            catch (Exception e)
            {
                UserDialogs.ShowMessage($"Помилка створення рапорту на премію: {e}", "Рапорт на премію",
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Сгенерувати рапорт на додаткову винагороду у Word документі на основі табеля у активному ЄЖООС файлі.
        /// </summary>
        public void OnAddRewardClicked(IRibbonControl control)
        {
            Application excelApp = (Application)ExcelDnaUtil.Application;
            try
            {
                new MoneyReportWorkflow(excelApp).CreateAddRewardReport(excelApp.ActiveWorkbook);
            }
            catch (Exception e)
            {
                UserDialogs.ShowMessage($"Помилка створення рапорту на ГЗ, рапорт не створено. {e}", "Рапорт на ГЗ",
                    MessageBoxIcon.Error);
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
        /// Заповнює Залученість (Job) у Табелі Залученості активної книги на основі обраної Стройової Записки.
        /// </summary>
        public void OnGetTabelZaluchenostiClicked(IRibbonControl control)
        {
            Application excelApp = (Application)ExcelDnaUtil.Application;
            try
            {
                new TabelZaluchenostiWorkflow(excelApp).FillJobsFromStroyova(excelApp.ActiveWorkbook);
            }
            catch (Exception e)
            {
                UserDialogs.ShowMessage($"Помилка при завантаженні табеля залученості: {e}", "Табель залученості");
            }
        }
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