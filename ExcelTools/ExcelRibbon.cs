using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ZsuTools.GenerateTabel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace ZsuTools
{
    [ComVisible( true )]
    public class ZsuToolsRibbon : ExcelRibbon
    {
        private IRibbonUI _ribbon;

        // Return the ribbon XML for Excel to render
        public override string GetCustomUI( string ribbonId )
        {
            var customUI =  @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>"
                
#if DEBUG
            +"<tab id='tabZsuTools_DEBUG' label='ЗСУ debug'>"
#else
            +"<tab id='tabZsuTools' label='ЗСУ'>"
#endif
                   + 
      @"<group id='groupMain' label='ЄЖООС'>
          <button id='btnAddReward' label='Рапорт на додаткову винагороду' size='large' imageMso='InternationalCurrency' showImage='true' onAction='OnAddRewardClicked' />
        </group>
        <group id='groupStroyova' label='Стройова'>
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
        public void OnLoad( IRibbonUI ribbonUi )
        {
            _ribbon = ribbonUi;
        }

        public void OnAddRewardClicked( IRibbonControl control )
        {
            Application excelApp = (Application) ExcelDnaUtil.Application;
            try
            {
                var workBook = excelApp.ActiveWorkbook;
                //excelApp.DisplayStatusBar = true;

                void UpdateStatusFunc( String status )
                {
                    UpdateStatus( excelApp, status );
                }

                MoneyReportGenerator.CreateMoney_100_30_Report(workBook, UpdateStatusFunc);
            }
            catch ( Exception e )
            {
                MessageBox.Show( $"Error catched, money report is discarded. {e}", "MoneyReport", MessageBoxButtons.OK,
                        MessageBoxIcon.Error );
            }
            finally
            {
                excelApp.StatusBar        = false; 
                //excelApp.DisplayStatusBar = false;
            }
        }

        public void OnGetTabelFromStroyovaClicked( IRibbonControl control )
        {
            Application excelApp = (Application) ExcelDnaUtil.Application;
            try
            {
                var workBook = excelApp.ActiveWorkbook;
                //excelApp.DisplayStatusBar = true;

                var stroyova = new СтройоваЗаписка(workBook);
            }
            catch ( Exception e )
            {
                MessageBox.Show( $"Виникла помилка, створення табеля зі Стройової записки перервано. {e}", "Стройова записка", MessageBoxButtons.OK,
                        MessageBoxIcon.Error );
            }
            finally
            {
                excelApp.StatusBar        = false; 
                //excelApp.DisplayStatusBar = false;
            }
        }
        
        public void OnGetTabelFromStroyovaBatchClicked( IRibbonControl control )
        {
            Application excelApp = (Application) ExcelDnaUtil.Application;
            try
            {
                var workBook = excelApp.ActiveWorkbook;
                //excelApp.DisplayStatusBar = true;

                var batchResult = BatchProcessor.ProcessMonthlyFiles();
                ResultTableGenerator.GenerateSummaryTable(batchResult);
                MessageBox.Show($"Табель зі Стройової Записки сформовано успішно, оброблено {batchResult.Count} файлів.", "Стройова записка", MessageBoxButtons.OK, MessageBoxIcon.Information);        
            }
            catch ( Exception e )
            {
                MessageBox.Show( $"Виникла помилка, створення табеля зі Стройової записки перервано. {e}", "Стройова записка", MessageBoxButtons.OK,
                    MessageBoxIcon.Error );
            }
            finally
            {
                excelApp.StatusBar        = false; 
                //excelApp.DisplayStatusBar = false;
            }
        }
        

        private void UpdateStatus( Application excelApp, string operationName )
        {
            excelApp.StatusBar = $"{operationName} {UpdateStates[_updateStateIndex++ % UpdateStates.Length]}";
        }

        private static readonly string[]    UpdateStates      = new []{"-", "\\", "|", "/"};
        private                 int         _updateStateIndex = 0;


    }
}