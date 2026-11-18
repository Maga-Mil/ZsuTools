using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using ZsuTools.Tables;

namespace ZsuTools.Tests
{
    /// <summary>
    /// Ручні тести (не unit-тести): запускаються вручну з Excel, потребують відкритого Табеля Залученості.
    /// </summary>
    public static class TabelZaluchenostiTests
    {
        private const int MaxReportedDifferences = 30;

        /// <summary>
        /// Читає Табель Залученості через COM з поточної (активної) книги, зберігає її копію у тимчасовий файл,
        /// читає копію через ClosedXML і порівнює два результати. Підсумок показує у MessageBox.
        /// </summary>
        public static void CompareComAndClosedXml()
        {
            var excelApp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
            var workbook = excelApp.ActiveWorkbook;
            if (workbook == null)
            {
                MessageBox.Show("Немає відкритої книги. Відкрийте Табель Залученості в Excel.",
                    "Тест Табеля Залученості");
                return;
            }

            // Файл відкритий в Excel, тому діалог його не дозволяє обрати. Робимо знімок поточного стану книги
            // (разом з незбереженими змінами) у тимчасовий файл і читаємо його через ClosedXML.
            string extension = Path.GetExtension(workbook.FullName);
            if (string.IsNullOrEmpty(extension))
            {
                extension = ".xlsx";
            }

            string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + extension);

            TabelZaluchenosti comTabel;
            TabelZaluchenosti closedXmlTabel;
            try
            {
                comTabel = new TabelZaluchenosti(workbook);

                workbook.SaveCopyAs(tempPath);
                using (var stream = new FileStream(tempPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    closedXmlTabel = new TabelZaluchenosti(stream);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Помилка читання Табеля Залученості: {ex.Message}", "Тест Табеля Залученості",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            finally
            {
                try
                {
                    File.Delete(tempPath);
                }
                catch (IOException)
                {
                }
            }

            var differences = Compare(comTabel, closedXmlTabel);

            var report = new StringBuilder();
            report.AppendLine($"COM: книга '{workbook.Name}', записів: {comTabel.Items.Count}");
            report.AppendLine($"ClosedXML: знімок книги '{workbook.Name}', записів: {closedXmlTabel.Items.Count}");
            report.AppendLine();

            if (differences.Count == 0)
            {
                report.AppendLine("Розбіжностей немає: обидва способи прочитали однакові дані.");
            }
            else
            {
                report.AppendLine($"Знайдено розбіжностей: {differences.Count}");
                report.AppendLine();
                for (int i = 0; i < Math.Min(differences.Count, MaxReportedDifferences); i++)
                {
                    report.AppendLine(differences[i]);
                }

                if (differences.Count > MaxReportedDifferences)
                {
                    report.AppendLine($"... і ще {differences.Count - MaxReportedDifferences}");
                }
            }

            MessageBox.Show(report.ToString(), "Тест Табеля Залученості", MessageBoxButtons.OK,
                differences.Count == 0 ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
        }

        /// <summary>
        /// Порівнює два Табелі Залученості (реєстри значень і записи). Повертає список розбіжностей.
        /// </summary>
        public static List<string> Compare(TabelZaluchenosti expected, TabelZaluchenosti actual)
        {
            var differences = new List<string>();

            CompareRegistries(expected, actual, differences);

            if (expected.Items.Count != actual.Items.Count)
            {
                differences.Add($"Кількість записів: COM = {expected.Items.Count}, ClosedXML = {actual.Items.Count}");
            }

            int count = Math.Min(expected.Items.Count, actual.Items.Count);
            for (int i = 0; i < count; i++)
            {
                CompareItems(i, expected.Items[i], actual.Items[i], differences);
            }

            return differences;
        }

        private static void CompareRegistries(TabelZaluchenosti expected, TabelZaluchenosti actual,
            List<string> differences)
        {
            for (int i = 0; i < expected.Jobs.AllValues.Count; i++)
            {
                var e = expected.Jobs.AllValues[i];
                var a = actual.Jobs.AllValues[i];
                if (e.Code != a.Code || e.Value != a.Value)
                {
                    differences.Add($"Jobs[{e.Code}]: COM = '{e.Value}', ClosedXML = '{a.Value}' ({a.Code})");
                }
            }

            for (int i = 0; i < expected.Presence.AllValues.Count; i++)
            {
                var e = expected.Presence.AllValues[i];
                var a = actual.Presence.AllValues[i];
                if (e.Value != a.Value)
                {
                    differences.Add($"Presence[{i}]: COM = '{e.Value}', ClosedXML = '{a.Value}'");
                }
            }
        }

        private static void CompareItems(int index, TabelZaluchenosti.Item e, TabelZaluchenosti.Item a,
            List<string> differences)
        {
            // Номер рядка у Табелі: 2 заголовки + 1-based індекс
            string row = $"Рядок {index + 3} ({e.PersonFullName})";

            Check(differences, row, nameof(e.LineNumber), e.LineNumber, a.LineNumber);
            Check(differences, row, nameof(e.UnitCode), e.UnitCode, a.UnitCode);
            Check(differences, row, nameof(e.UnitName), e.UnitName, a.UnitName);
            Check(differences, row, nameof(e.PositionId), e.PositionId, a.PositionId);
            Check(differences, row, nameof(e.PositionName), e.PositionName, a.PositionName);
            Check(differences, row, nameof(e.PersonId), e.PersonId, a.PersonId);
            Check(differences, row, nameof(e.PersonRank), e.PersonRank, a.PersonRank);
            Check(differences, row, nameof(e.PersonFullName), e.PersonFullName, a.PersonFullName);
            Check(differences, row, nameof(e.PersonTaxCode), e.PersonTaxCode, a.PersonTaxCode);
            Check(differences, row, nameof(e.Presence), e.Presence?.Value, a.Presence?.Value);
            Check(differences, row, nameof(e.Job), e.Job?.Code, a.Job?.Code);
            Check(differences, row, nameof(e.JobDetails), e.JobDetails, a.JobDetails);
            Check(differences, row, nameof(e.JobStartDate), e.JobStartDate, a.JobStartDate);
            Check(differences, row, nameof(e.JobStartTime), e.JobStartTime, a.JobStartTime);
            Check(differences, row, nameof(e.JobFinishDate), e.JobFinishDate, a.JobFinishDate);
            Check(differences, row, nameof(e.JobFinishTime), e.JobFinishTime, a.JobFinishTime);
        }

        private static void Check<T>(List<string> differences, string row, string field, T expected, T actual)
        {
            if (!EqualityComparer<T>.Default.Equals(expected, actual))
            {
                differences.Add($"{row}, {field}: COM = '{expected}', ClosedXML = '{actual}'");
            }
        }
    }
}