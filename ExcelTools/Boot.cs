using ExcelDna.Integration;
using ExcelDna.IntelliSense;

namespace ZsuTools
{
    public class ZsuToolsAddin : IExcelAddIn
    {
        public void AutoOpen()
        {
            // Вмикаємо живі підказки для всіх функцій аддону
            IntelliSenseServer.Install();
        }

        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
        }
    }
}