using System.IO;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;

namespace ZsuTools.UI
{
    /// <summary>
    /// Діалоги вибору файлів та повідомлення користувачу, прив'язані до головного вікна Excel.
    /// </summary>
    public static class UserDialogs
    {
        public const string ExcelFilter = "Файли Excel (*.xlsx)|*.xlsx|Усі файли (*.*)|*.*";

        /// <summary>
        /// Показує діалог вибору файлу. Повертає false, якщо користувач скасував вибір.
        /// </summary>
        public static bool TryOpenFile(string filter, string title, out string fileName)
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = filter;
                openFileDialog.Title = title;

                if (openFileDialog.ShowDialog(ExcelWindow.Instance) != DialogResult.OK)
                {
                    fileName = null;
                    return false;
                }

                fileName = openFileDialog.FileName;
                return true;
            }
        }

        /// <summary>
        /// Показує діалог вибору файлу і відкриває обрану книгу через COM лише для читання.
        /// Закрити книгу — відповідальність викликача.
        /// </summary>
        public static bool TryOpenWorkbook(Application excelApp, string filter, string title, out Workbook workbook)
        {
            if (!TryOpenFile(filter, title, out var fileName))
            {
                workbook = null;
                return false;
            }

            workbook = excelApp.Workbooks.Open(fileName, ReadOnly: true);
            return true;
        }

        /// <summary>
        /// Показує діалог вибору файлу і відкриває його як потік для читання.
        /// Звільнити потік — відповідальність викликача.
        /// </summary>
        public static bool TryOpenStream(string filter, string title, out Stream stream)
        {
            if (!TryOpenFile(filter, title, out var fileName))
            {
                stream = null;
                return false;
            }

            stream = File.OpenRead(fileName);
            return true;
        }

        public static void ShowMessage(string text, string caption, MessageBoxIcon icon = MessageBoxIcon.None)
        {
            MessageBox.Show(ExcelWindow.Instance, text, caption, MessageBoxButtons.OK, icon);
        }

        /// <summary>
        /// Показує повідомлення з кнопками OK / Скасувати. Повертає true, якщо натиснуто OK.
        /// За замовчуванням фокус на «Скасувати», щоб випадковий Enter нічого не змінив.
        /// </summary>
        public static bool Confirm(string text, string caption, MessageBoxIcon icon = MessageBoxIcon.Warning)
        {
            return MessageBox.Show(ExcelWindow.Instance, text, caption, MessageBoxButtons.OKCancel, icon,
                MessageBoxDefaultButton.Button2) == DialogResult.OK;
        }
    }
}