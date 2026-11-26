using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using ZsuTools.Tables;
using ZsuTools.UI;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace ZsuTools.Workflows
{
    /// <summary>
    /// Процеси роботи з Табелем Залученості.
    /// </summary>
    public class TabelZaluchenostiWorkflow
    {
        public TabelZaluchenostiWorkflow(Application excelApp)
        {
            this.excelApp = excelApp;
        }

        /// <summary>
        /// Заповнює Залученість (Job) у Табелі Залученості на основі обраної Стройової Записки:
        /// БЧ → BATTALION_DEFENSE_SECTOR, РЗ → COMBAT_SUPPORT, решта — за назвою місцезнаходження.
        /// </summary>
        public void FillJobsFromStroyova(Workbook tabelWorkbook)
        {
            if (!TryLoadTabel(tabelWorkbook, out var tz))
            {
                return;
            }

            if (!TryLoadStroyova(tabelWorkbook, out var sz))
            {
                return;
            }

            if (!ValidateSameRecords(tz, sz))
            {
                return;
            }

            if (!TryBuildAssignments(tz, sz, out var assignments))
            {
                return;
            }

            if (!ConfirmPresenceJobWarnings(tz, assignments))
            {
                return;
            }

            ApplyAndSave(tz, sz, assignments);
        }

        private const string Caption = "Табель залученості";
        private const int MaxWarningsInMessage = 25;

        private readonly Application excelApp;

        private static bool TryLoadTabel(Workbook workBook, out TabelZaluchenosti tz)
        {
            tz = null;
            if (workBook == null)
            {
                UserDialogs.ShowMessage("Не завантажено Табель Залученості.", Caption);
                return false;
            }

            var ws = ExcelUtils.FindWorksheetByNameContains(workBook, "Табелювання");
            if (ws == null)
            {
                UserDialogs.ShowMessage("Не знайдено таблицю 'Табелювання', це не схоже на Табель Залученості",
                    Caption);
                return false;
            }

            tz = new TabelZaluchenosti(workBook);
            return true;
        }

        private bool TryLoadStroyova(Workbook tabelWorkbook, out StroyovaZapyska sz)
        {
            sz = null;
            if (!UserDialogs.TryOpenWorkbook(excelApp, UserDialogs.ExcelFilter,
                    "Оберіть Стройову Записку, щоб скопіювати дані до Табеля Залученості", out var szWb))
            {
                return false;
            }

            try
            {
                sz = new StroyovaZapyska(szWb);
            }
            finally
            {
                szWb.Close(false);
            }

            // Табель міг перестати бути активною книгою, поки відкривалася Стройова
            tabelWorkbook.Activate();
            return true;
        }

        private static bool ValidateSameRecords(TabelZaluchenosti tz, StroyovaZapyska sz)
        {
            if (tz.Items.Count != sz.Items.Count)
            {
                UserDialogs.ShowMessage(
                    $"Кількість записів у Табелі Залученості ({tz.Items.Count}) не збігається з кількістю записів у Стройовій Записці ({sz.Items.Count}). Операця буде зупинена.",
                    Caption);
                return false;
            }

            var mismatchedRecords = new List<string>();
            foreach (var tzItem in tz.Items)
            {
                var person = sz.Items.SingleOrDefault(r =>
                    r.Person.FullName.Equals(tzItem.PersonFullName, StringComparison.InvariantCultureIgnoreCase));
                if (person == null)
                {
                    mismatchedRecords.Add(tzItem.PersonFullName);
                }
            }

            if (mismatchedRecords.Count > 0)
            {
                UserDialogs.ShowMessage(
                    $"У Табелі Залученості присутні записи, які відсутні у Стройовій Записці:\n\r {String.Join(", ", mismatchedRecords)}. \n\rТреба виправити ПІБ у Стройовій Записці відповідно до Табеля Залученості. Операця буде зупинена.",
                    Caption);
                return false;
            }

            return true;
        }

        /// <summary>
        /// Для кожної особи зі Стройової знаходить відповідну особу в Табелі (за ПІБ) і значення Залученості.
        /// Якщо є хоч одна проблема — показує зведене повідомлення і повертає false.
        /// </summary>
        private static bool TryBuildAssignments(TabelZaluchenosti tz, StroyovaZapyska sz,
            out List<KeyValuePair<TabelZaluchenosti.Item, TabelZaluchenosti.JobValue>> assignments)
        {
            var tzByName = tz.Items.ToLookup(i => i.PersonFullName, StringComparer.InvariantCultureIgnoreCase);
            var notFound = new List<string>();
            var ambiguous = new List<string>();
            var unknownLocations = new List<string>();
            var missingJobValues = new List<string>();
            assignments = new List<KeyValuePair<TabelZaluchenosti.Item, TabelZaluchenosti.JobValue>>();

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

            if (notFound.Count == 0 && ambiguous.Count == 0 && unknownLocations.Count == 0 &&
                missingJobValues.Count == 0)
            {
                return true;
            }

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
            UserDialogs.ShowMessage(message.ToString(), Caption);
            return false;
        }

        /// <summary>
        /// Неузгодженість Наявності і Залученості — не помилка, а попередження: показує перелік осіб
        /// і питає, чи застосовувати зміни. Повертає false, якщо користувач скасував операцію.
        /// </summary>
        private static bool ConfirmPresenceJobWarnings(TabelZaluchenosti tz,
            List<KeyValuePair<TabelZaluchenosti.Item, TabelZaluchenosti.JobValue>> assignments)
        {
            var proposedJobs = assignments.ToDictionary(a => a.Key, a => a.Value);
            var warnings = GetPresenceJobWarnings(tz, proposedJobs);
            if (warnings.Count == 0)
            {
                return true;
            }

            var message = new StringBuilder();
            message.AppendLine($"Наявність не узгоджується із Залученістю ({warnings.Count}):");
            message.AppendLine();
            // MessageBox не прокручується, тому довгий перелік обрізаємо
            foreach (var warning in warnings.Take(MaxWarningsInMessage))
            {
                message.AppendLine(warning);
            }

            if (warnings.Count > MaxWarningsInMessage)
            {
                message.AppendLine($"…та ще {warnings.Count - MaxWarningsInMessage}");
            }

            message.AppendLine();
            message.Append("OK — заповнити Табель Залученості попри попередження.\n" +
                           "Скасувати — зупинити операцію, Табель Залученості не буде змінено.");

            return UserDialogs.Confirm(message.ToString(), Caption);
        }

        /// <summary>
        /// Перевіряє узгодженість Наявності (Presence) і Залученості (Job) для всіх записів Табеля:
        /// PRESENT → BATTALION_DEFENSE_SECTOR або COMBAT_SUPPORT; відпустки (VACATION_*) → VACATION;
        /// відрядження (TRIP_*) → TRIP; AWOL → AWOL.
        /// Для записів із <paramref name="proposedJobs"/> перевіряється нове (ще не записане) значення Job.
        /// Повертає текстовий перелік порушень (порожній, якщо все узгоджено).
        /// </summary>
        private static List<string> GetPresenceJobWarnings(TabelZaluchenosti tz,
            IReadOnlyDictionary<TabelZaluchenosti.Item, TabelZaluchenosti.JobValue> proposedJobs)
        {
            var presences = tz.Presence;
            var jobs = tz.Jobs;
            var warnings = new List<string>();

            foreach (var item in tz.Items)
            {
                var presence = item.Presence;
                if (!proposedJobs.TryGetValue(item, out var job))
                {
                    job = item.Job;
                }

                job = job ?? TabelZaluchenosti.JobValue.Empty;
                TabelZaluchenosti.JobValue[] allowedJobs;

                if (ReferenceEquals(presence, presences.PRESENT))
                {
                    allowedJobs = new[] { jobs.BATTALION_DEFENSE_SECTOR, jobs.COMBAT_SUPPORT };
                }
                else if (ReferenceEquals(presence, presences.VACATION_ANNUAL) ||
                         ReferenceEquals(presence, presences.VACATION_FAMILY_REASONS) ||
                         ReferenceEquals(presence, presences.VACATION_MEDICAL) ||
                         ReferenceEquals(presence, presences.VACATION_MEDICAL_SICK))
                {
                    allowedJobs = new[] { jobs.VACATION };
                }
                else if (ReferenceEquals(presence, presences.TRIP_FOR_PERIOD) ||
                         ReferenceEquals(presence, presences.TRIP_UNTIL_DIRECTIVE) ||
                         ReferenceEquals(presence, presences.TRIP_FOR_EDUCATION))
                {
                    allowedJobs = new[] { jobs.TRIP };
                }
                else if (ReferenceEquals(presence, presences.AWOL))
                {
                    allowedJobs = new[] { jobs.AWOL };
                }
                else
                {
                    continue; // Для інших значень Наявності правил поки немає
                }

                if (!allowedJobs.Contains(job))
                {
                    var jobText = string.IsNullOrEmpty(job.Value) ? "(порожньо)" : job.Value;
                    warnings.Add($"{item.PersonFullName}: Наявність «{presence.Value}», Залученість «{jobText}»");
                }
            }

            return warnings;
        }

        private static void ApplyAndSave(TabelZaluchenosti tz, StroyovaZapyska sz,
            List<KeyValuePair<TabelZaluchenosti.Item, TabelZaluchenosti.JobValue>> assignments)
        {
            foreach (var assignment in assignments)
            {
                assignment.Key.Job = assignment.Value;
            }

            var written = tz.Save();

            UserDialogs.ShowMessage(
                $"Залученість заповнено зі Стройової Записки на {sz.Date:dd.MM.yyyy}.\n" +
                $"Оброблено осіб: {assignments.Count}, змінено комірок: {written}.",
                Caption, MessageBoxIcon.Information);
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
    }
}