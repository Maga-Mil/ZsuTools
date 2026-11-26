using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;

namespace ZsuTools.Tables
{
    public class TabelZaluchenosti
    {
        /// <summary>
        /// Записи Табеля. Структурні зміни (додавання/видалення) не підтримуються; змінювати можна лише
        /// поля записів (Job, JobDetails), після чого викликати <see cref="Save"/>.
        /// </summary>
        public IReadOnlyList<Item> Items => items;

        public JobValue.Registry Jobs { get; private set; }
        public PresenceValue.Registry Presence { get; private set; }

        /// <summary>
        /// Чи є зміни Job / JobDetails відносно прочитаного (або останнього збереженого) стану.
        /// </summary>
        public bool HasChanges => GetChanges().Count > 0;

        public TabelZaluchenosti(Workbook wb)
        {
            var jobsValues = GetDictionaryFromNamedRanges(wb, "InvVal[val_code]", "InvVal[val_name]");
            Jobs = new JobValue.Registry(jobsValues);

            var presenceValues = GetDictionaryFromNamedRanges(wb, "PersonVal[val_code]", "PersonVal[val_name]");
            Presence = new PresenceValue.Registry(presenceValues);

            var mainWorkSheet = ExcelUtils.FindWorksheetByNameContains(wb, "Табелювання");
            items.AddRange(ReadTableFromWorksheet(mainWorkSheet, Jobs, Presence));

            comSheet = mainWorkSheet;
            TakeSnapshot();
        }

        /// <summary>
        /// Створює Табель Залученості з Excel файлу, використовуючи ClosedXML (без Excel Interop).
        /// </summary>
        public TabelZaluchenosti(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                throw new ArgumentException("Передано некоректний шлях до файлу.", nameof(filePath));
            }

            using (var workbook = new XLWorkbook(filePath))
            {
                Load(workbook);
            }
        }

        /// <summary>
        /// Створює Табель Залученості з Stream Excel файлу, використовуючи ClosedXML (без Excel Interop).
        /// </summary>
        public TabelZaluchenosti(Stream stream)
        {
            if (stream == null)
            {
                throw new ArgumentException("Передано некоректний потік даних.", nameof(stream));
            }

            using (var workbook = new XLWorkbook(stream))
            {
                Load(workbook);
            }
        }

        /// <summary>
        /// Записує у аркуш «Табелювання» відкритої книги лише ті комірки Job / JobDetails, що змінилися.
        /// Книга на диск не зберігається.
        /// </summary>
        /// <returns>Кількість записаних комірок.</returns>
        public int Save()
        {
            if (comSheet == null)
            {
                throw new NotSupportedException(
                    "Збереження підтримується лише для Табеля Залученості, відкритого в Excel.");
            }

            var changes = GetChanges();
            if (changes.Count == 0)
            {
                return 0;
            }

            // Рядок шукаємо за ID особи в момент збереження: між читанням і збереженням рядки на аркуші
            // могли вставити/видалити/пересортувати (користувач або інший макрос).
            // Спершу визначаємо рядки для всіх змін, щоб не було часткового запису.
            var rowsById = BuildRowIndex(comSheet);
            var targets = new List<KeyValuePair<int, CellChange>>(changes.Count);
            foreach (var change in changes)
            {
                if (!rowsById.TryGetValue(change.PersonId, out int row))
                {
                    throw new InvalidOperationException(
                        $"Не знайдено рядок особи '{change.PersonFullName}' на аркуші (можливо, рядок видалено " +
                        "після читання). Перечитайте Табель Залученості та повторіть операцію.");
                }

                targets.Add(new KeyValuePair<int, CellChange>(row, change));
            }

            var app = comSheet.Application;
            bool screenUpdating = app.ScreenUpdating;
            app.ScreenUpdating = false;
            try
            {
                foreach (var target in targets)
                {
                    ((Range)comSheet.Cells[target.Key, target.Value.Column]).Value2 = target.Value.Value;
                }
            }
            finally
            {
                app.ScreenUpdating = screenUpdating;
            }

            TakeSnapshot();
            return changes.Count;
        }

        public class Item
        {
            internal Item(Guid personId)
            {
                PersonId = personId;
            }

            public int LineNumber { get; set; }
            public int UnitCode { get; set; }
            public string UnitName { get; set; }
            public int PositionId { get; set; }
            public string PositionName { get; set; }

            /// <summary>
            /// ID особи — ключ для пошуку рядка на аркуші під час збереження, тому змінювати його не можна.
            /// </summary>
            public Guid PersonId { get; }

            public string PersonRank { get; set; }
            public string PersonFullName { get; set; }
            public ulong PersonTaxCode { get; set; }
            public PresenceValue Presence { get; set; }
            public JobValue Job { get; set; }
            public string JobDetails { get; set; }
            public DateTimeOffset? JobStartDate { get; set; }
            public TimeSpan? JobStartTime { get; set; }
            public DateTimeOffset? JobFinishDate { get; set; }
            public TimeSpan? JobFinishTime { get; set; }

            internal Item Clone() => (Item)MemberwiseClone();
        }

        public sealed class JobValue
        {
            // Приватний конструктор — створювати екземпляри може тільки сам клас/фабрика
            private JobValue(string code, string value)
            {
                Code = code;
                Value = value;
            }

            /// <summary>
            /// Внутрішній код/ідентифікатор елемента (наприклад, "BATTALION_DEFENSE_SECTOR")
            /// </summary>
            public string Code { get; }

            /// <summary>
            /// Локалізоване або зчитане з Excel значення
            /// </summary>
            public string Value { get; }

            // Константні/дефолтні екземпляри (статичні, тому не викликають рекурсію екземплярів)
            public static readonly JobValue Empty = new JobValue(nameof(Empty), string.Empty);
            public static readonly JobValue Unknown = new JobValue(nameof(Unknown), "Unknown");

            public override string ToString() => Value;

            public class Registry
            {
                public JobValue BATTALION_DEFENSE_SECTOR { get; }
                public JobValue COMBAT_SUPPORT { get; }
                public JobValue VACATION { get; }
                public JobValue SHORT_TERM_TREATMENT { get; }
                public JobValue LONG_TERM_TREATMENT { get; }
                public JobValue TRIP { get; }
                public JobValue AWOL { get; }

                public JobValue Empty => JobValue.Empty;
                public JobValue Unknown => JobValue.Unknown;

                public IReadOnlyList<JobValue> AllValues { get; }

                public Registry(IReadOnlyDictionary<string, string> values)
                {
                    BATTALION_DEFENSE_SECTOR = Create(values, nameof(BATTALION_DEFENSE_SECTOR));
                    COMBAT_SUPPORT = Create(values, nameof(COMBAT_SUPPORT));
                    VACATION = Create(values, nameof(VACATION));
                    SHORT_TERM_TREATMENT = Create(values, nameof(SHORT_TERM_TREATMENT));
                    LONG_TERM_TREATMENT = Create(values, nameof(LONG_TERM_TREATMENT));
                    TRIP = Create(values, nameof(TRIP));
                    AWOL = Create(values, nameof(AWOL));

                    AllValues = new[]
                    {
                        BATTALION_DEFENSE_SECTOR,
                        COMBAT_SUPPORT,
                        VACATION,
                        SHORT_TERM_TREATMENT,
                        LONG_TERM_TREATMENT,
                        TRIP,
                        AWOL
                    };
                }

                private static JobValue Create(IReadOnlyDictionary<string, string> dict, string key)
                {
                    return dict.TryGetValue(key, out var val)
                        ? new JobValue(key, val)
                        : JobValue.Unknown;
                }
            }
        }

        public class PresenceValue
        {
            private PresenceValue(string value)
            {
                Value = value;
            }

            public string Value { get; }

            public override string ToString()
            {
                return Value;
            }

            public class Registry
            {
                /// <summary>
                /// В наявності
                /// </summary>
                public readonly PresenceValue PRESENT;

                /// <summary>
                /// Відпустка щорічна основна
                /// </summary>
                public readonly PresenceValue VACATION_ANNUAL;

                /// <summary>
                /// Відпустка за сімейними обставинами
                /// </summary>
                public readonly PresenceValue VACATION_FAMILY_REASONS;

                /// <summary>
                /// Відпустка на лікування після поранення, травми, контузії
                /// </summary>
                public readonly PresenceValue VACATION_MEDICAL;

                /// <summary>
                /// Відпустка на лікування у зв’язку з хворобою
                /// </summary>
                public readonly PresenceValue VACATION_MEDICAL_SICK;

                /// <summary>
                /// Відрядження строкове
                /// </summary>
                public readonly PresenceValue TRIP_FOR_PERIOD;

                /// <summary>
                /// Відрядження до розпорядження
                /// </summary>
                public readonly PresenceValue TRIP_UNTIL_DIRECTIVE;

                /// <summary>
                /// Відрядження на навчання в НЦ, ВВНЗ
                /// </summary>
                public readonly PresenceValue TRIP_FOR_EDUCATION;

                /// <summary>
                /// СЗЧ
                /// </summary>
                public readonly PresenceValue AWOL;

                public readonly PresenceValue Empty = new PresenceValue(string.Empty);
                public readonly PresenceValue Unknown = new PresenceValue("Unknown");

                public readonly IReadOnlyList<PresenceValue> AllValues;

                public Registry(IReadOnlyDictionary<string, string> allValues)
                {
                    PRESENT = new PresenceValue(allValues[nameof(PRESENT)]);
                    VACATION_ANNUAL = new PresenceValue(allValues[nameof(VACATION_ANNUAL)]);
                    VACATION_FAMILY_REASONS = new PresenceValue(allValues[nameof(VACATION_FAMILY_REASONS)]);
                    VACATION_MEDICAL = new PresenceValue(allValues[nameof(VACATION_MEDICAL)]);
                    VACATION_MEDICAL_SICK = new PresenceValue(allValues[nameof(VACATION_MEDICAL_SICK)]);
                    TRIP_FOR_PERIOD = new PresenceValue(allValues[nameof(TRIP_FOR_PERIOD)]);
                    TRIP_UNTIL_DIRECTIVE = new PresenceValue(allValues[nameof(TRIP_UNTIL_DIRECTIVE)]);
                    TRIP_FOR_EDUCATION = new PresenceValue(allValues[nameof(TRIP_FOR_EDUCATION)]);
                    AWOL = new PresenceValue(allValues[nameof(AWOL)]);

                    AllValues = new[]
                    {
                        PRESENT,
                        VACATION_ANNUAL,
                        VACATION_FAMILY_REASONS,
                        VACATION_MEDICAL,
                        VACATION_MEDICAL_SICK,
                        TRIP_FOR_PERIOD,
                        TRIP_UNTIL_DIRECTIVE,
                        TRIP_FOR_EDUCATION,
                        AWOL
                    };
                }
            }
        }

        // Номери колонок аркуша «Табелювання» (1-based)
        private const int ColPersonId = 6;
        private const int ColJob = 12;
        private const int ColJobDetails = 13;

        private readonly List<Item> items = new List<Item>();

        /// <summary>
        /// Знімок записів на момент читання (або останнього збереження): snapshot[i] — копія items[i].
        /// </summary>
        private List<Item> snapshot = new List<Item>();

        /// <summary>
        /// Аркуш, з якого прочитано дані через COM. Для ClosedXML-джерел — null (збереження поки не підтримується).
        /// </summary>
        private readonly Worksheet comSheet;

        private void Load(XLWorkbook workbook)
        {
            var jobsValues = GetDictionaryFromNamedRanges(workbook, "InvVal[val_code]", "InvVal[val_name]");
            Jobs = new JobValue.Registry(jobsValues);

            var presenceValues = GetDictionaryFromNamedRanges(workbook, "PersonVal[val_code]", "PersonVal[val_name]");
            Presence = new PresenceValue.Registry(presenceValues);

            var mainWorkSheet = workbook.Worksheets.FirstOrDefault(ws =>
                ws.Name.IndexOf("Табелювання", StringComparison.InvariantCultureIgnoreCase) >= 0);
            if (mainWorkSheet == null)
            {
                throw new ArgumentException("Не знайдено аркуш 'Табелювання', це не схоже на Табель Залученості.");
            }

            items.AddRange(ReadTableFromWorksheet(mainWorkSheet, Jobs, Presence));
            TakeSnapshot();
        }

        #region Changes tracking / Save

        /// <summary>
        /// Будує відповідність «ID особи → номер рядка аркуша» за тими ж межами, що й читання:
        /// з рядка 3 до першої порожньої комірки в колонці A. Читає дані одним COM-викликом.
        /// </summary>
        private static Dictionary<Guid, int> BuildRowIndex(Worksheet sheet)
        {
            var result = new Dictionary<Guid, int>();

            int totalRows = sheet.UsedRange.Rows.Count;
            if (totalRows < 3)
            {
                return result;
            }

            var data = (object[,])sheet.Range[sheet.Cells[1, 1], sheet.Cells[totalRows, ColPersonId]].Value2;
            for (int row = 3; row <= totalRows; row++)
            {
                if (IsCellEmpty(data[row, 1]))
                {
                    break;
                }

                var id = ParseGuid(data[row, ColPersonId]);
                if (result.TryGetValue(id, out var firstRow))
                {
                    throw new InvalidOperationException(
                        $"ID особи у рядку {row} аркуша повторюється (уперше — у рядку {firstRow}). " +
                        "Виправте Табель Залученості та повторіть операцію.");
                }

                result.Add(id, row);
            }

            return result;
        }

        private void TakeSnapshot()
        {
            snapshot = items.Select(i => i.Clone()).ToList();
        }

        /// <summary>
        /// Обчислює список змінених комірок (Job / JobDetails). Не залежить від способу запису (COM / файл).
        /// </summary>
        private List<CellChange> GetChanges()
        {
            var result = new List<CellChange>();

            for (var i = 0; i < items.Count; i++)
            {
                var item = items[i];
                var original = snapshot[i];

                var job = item.Job ?? JobValue.Empty;
                var originalJob = original.Job ?? JobValue.Empty;
                if (job.Code != originalJob.Code || job.Value != originalJob.Value)
                {
                    if (ReferenceEquals(job, JobValue.Unknown))
                    {
                        throw new InvalidOperationException(
                            $"Неможливо записати невідоме значення залученості для '{item.PersonFullName}'.");
                    }

                    object value = ReferenceEquals(job, JobValue.Empty) ? null : job.Value;
                    result.Add(new CellChange(original, ColJob, value));
                }

                // Обрізаємо пробіли так само, як читання (ParseString), щоб записане збігалося з перечитаним
                string details = (item.JobDetails ?? string.Empty).Trim();
                string originalDetails = (original.JobDetails ?? string.Empty).Trim();
                if (!string.Equals(details, originalDetails, StringComparison.Ordinal))
                {
                    object value = details.Length == 0 ? null : details;
                    result.Add(new CellChange(original, ColJobDetails, value));
                }
            }

            return result;
        }

        private sealed class CellChange
        {
            // ПІБ беремо зі знімка, а не з живого (змінюваного) запису
            public CellChange(Item original, int column, object value)
            {
                PersonId = original.PersonId;
                PersonFullName = original.PersonFullName;
                Column = column;
                Value = value;
            }

            public Guid PersonId { get; }
            public string PersonFullName { get; }
            public int Column { get; }
            public object Value { get; }
        }

        #endregion

        #region Reading: ClosedXML

        /// <summary>
        /// Читає рядки даних з аркуша через ClosedXML. Логіка ідентична COM-версії.
        /// </summary>
        private static List<Item> ReadTableFromWorksheet(IXLWorksheet sheet, JobValue.Registry jobValue,
            PresenceValue.Registry presenceValue)
        {
            if (sheet == null)
            {
                throw new ArgumentNullException(nameof(sheet), "Аркуш не може бути порожнім.");
            }

            var results = new List<Item>();
            int lastRow = sheet.LastRowUsed()?.RowNumber() ?? 0;

            // Рядки 1 і 2 — заголовки, дані починаються з рядка 3
            for (int row = 3; row <= lastRow; row++)
            {
                // Зупиняємось, щойно колонка A порожня
                if (IsCellEmpty(GetCellValue(sheet.Cell(row, 1))))
                {
                    break;
                }

                results.Add(ParseItem(col => GetCellValue(sheet.Cell(row, col)), jobValue, presenceValue));
            }

            return results;
        }

        /// <summary>
        /// Повертає значення комірки у тому ж представленні, що й COM Value2:
        /// числа і дати/час — double (OADate), текст — string, порожня комірка — null.
        /// </summary>
        private static object GetCellValue(IXLCell cell)
        {
            if (cell.IsMerged())
            {
                cell = cell.MergedRange().FirstCell();
            }

            XLCellValue value = cell.Value;

            if (value.IsBlank) return null;
            if (value.IsNumber) return value.GetNumber();
            if (value.IsDateTime) return value.GetDateTime().ToOADate();
            if (value.IsTimeSpan) return value.GetTimeSpan().TotalDays;
            if (value.IsBoolean) return value.GetBoolean();
            if (value.IsText) return value.GetText();

            // Помилки формул (#N/A тощо) вважаємо порожніми
            return null;
        }

        private static Dictionary<string, string> GetDictionaryFromNamedRanges(XLWorkbook workbook,
            string keyRangeName, string valRangeName)
        {
            var keys = ResolveColumnValues(workbook, keyRangeName);
            var values = ResolveColumnValues(workbook, valRangeName);

            var result = new Dictionary<string, string>();
            for (int i = 0; i < keys.Count; i++)
            {
                string key = keys[i];
                string val = i < values.Count ? values[i] : string.Empty;

                if (!string.IsNullOrEmpty(key) && !result.ContainsKey(key))
                {
                    result.Add(key, val);
                }
            }

            return result;
        }

        /// <summary>
        /// Повертає значення колонки за виразом виду "Table[column]" (розумна таблиця) або за іменем діапазону.
        /// </summary>
        private static List<string> ResolveColumnValues(XLWorkbook workbook, string rangeExpression)
        {
            var result = new List<string>();

            var match = Regex.Match(rangeExpression, @"^(?<table>[^\[]+)\[(?<column>[^\]]+)\]$");
            if (match.Success)
            {
                string tableName = match.Groups["table"].Value;
                string columnName = match.Groups["column"].Value;

                var table = workbook.Worksheets
                    .SelectMany(ws => ws.Tables)
                    .FirstOrDefault(t => t.Name.Equals(tableName, StringComparison.InvariantCultureIgnoreCase));
                if (table == null)
                {
                    throw new ArgumentException($"Не знайдено таблицю '{tableName}' у книзі.");
                }

                foreach (var cell in table.Field(columnName).DataCells)
                {
                    result.Add(cell.GetString().Trim());
                }

                return result;
            }

            if (workbook.DefinedNames.TryGetValue(rangeExpression, out var definedName))
            {
                foreach (var range in definedName.Ranges)
                {
                    foreach (var cell in range.Cells())
                    {
                        result.Add(cell.GetString().Trim());
                    }
                }

                return result;
            }

            throw new ArgumentException($"Не знайдено діапазон '{rangeExpression}' у книзі.");
        }

        #endregion

        #region Reading: COM

        /// <summary>
        /// Reads table data directly from the provided <see cref="Worksheet"/> instance parameter.
        /// </summary>
        /// <param name="sheet">The target Excel worksheet instance to read from.</param>
        /// <returns>List of parsed <see cref="Item"/> items.</returns>
        private static List<Item> ReadTableFromWorksheet(Worksheet sheet, JobValue.Registry jobValue,
            PresenceValue.Registry presenceValue)
        {
            if (sheet == null)
            {
                throw new ArgumentNullException(nameof(sheet), "Worksheet parameter cannot be null.");
            }

            var results = new List<Item>();

            // Find used range or targeted columns A to Q
            Range usedRange = sheet.UsedRange;
            if (usedRange == null)
            {
                return results;
            }

            int totalRows = usedRange.Rows.Count;

            // If less than 3 rows exist, there are no data rows (2 header rows skipped)
            if (totalRows < 3)
            {
                return results;
            }

            // Fetch range values into a 2D object array in a single COM call for max performance
            // Standard Excel COM array indexing is 1-based: data[row, column]
            object[,] data = (object[,])sheet.Range[sheet.Cells[1, 1], sheet.Cells[totalRows, 19]].Value2;

            // Row 1 and 2 are top header lines skipped. Start reading from Row 3.
            for (int row = 3; row <= totalRows; row++)
            {
                object colA = data[row, 1];

                // Stop execution as soon as Column A is empty or null
                if (IsCellEmpty(colA))
                {
                    break;
                }

                int currentRow = row;
                results.Add(ParseItem(col => data[currentRow, col], jobValue, presenceValue));
            }

            return results;
        }

        /// <summary>
        /// Зчитує дані з двох іменованих діапазонів у Dictionary через COM Interop для конкретної книги.
        /// Безпечно викликається з Ribbon UI та інших COM-подій.
        /// </summary>
        /// <param name="wb">Об'єкт робочої книги Excel (Workbook)</param>
        /// <param name="keyRangeName">Назва діапазону/колонки для ключів (наприклад, "PersonVal[val_code]")</param>
        /// <param name="valRangeName">Назва діапазону/колонки для значень (наприклад, "PersonVal[val_name]")</param>
        private static Dictionary<string, string> GetDictionaryFromNamedRanges(Workbook wb, string keyRangeName,
            string valRangeName)
        {
            var result = new Dictionary<string, string>();

            if (wb == null)
            {
                throw new ArgumentNullException(nameof(wb), "Передано порожній об'єкт Workbook.");
            }

            try
            {
                Range keyRange = ResolveRange(wb, keyRangeName);
                Range valRange = ResolveRange(wb, valRangeName);

                if (keyRange != null && valRange != null)
                {
                    object keysObj = keyRange.Value2;
                    object valsObj = valRange.Value2;

                    // Масив з кількох комірок (двовимірний 1-based масив у COM)
                    if (keysObj is object[,] keysArray && valsObj is object[,] valsArray)
                    {
                        int rowsCount = keysArray.GetLength(0);

                        for (int i = 1; i <= rowsCount; i++)
                        {
                            string key = keysArray[i, 1]?.ToString()?.Trim();
                            string val = valsArray[i, 1]?.ToString()?.Trim();

                            if (!string.IsNullOrEmpty(key) && !result.ContainsKey(key))
                            {
                                result.Add(key, val ?? string.Empty);
                            }
                        }
                    }
                    // Випадок, якщо в діапазоні лише 1 комірка
                    else if (keysObj != null)
                    {
                        string key = keysObj.ToString()?.Trim();
                        string val = valsObj?.ToString()?.Trim();

                        if (!string.IsNullOrEmpty(key))
                        {
                            result[key] = val ?? string.Empty;
                        }
                    }
                }
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"Помилка зчитування діапазонів у книзі '{wb.Name}': {ex.Message}");
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Загальна помилка: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// Отримує Range з книги: підтримує класичні Named Ranges та структурні посилання розумних таблиць.
        /// </summary>
        private static Range ResolveRange(Workbook wb, string rangeExpression)
        {
            var app = (Application)ExcelDnaUtil.Application;

            try
            {
                // 1. Пробуємо знайти як звичайне ім'я в wb.Names
                Name nameObj = wb.Names.Item(rangeExpression);
                return nameObj.RefersToRange;
            }
            catch
            {
                // 2. Якщо це структурне посилання (наприклад, "PersonVal[val_code]"), 
                // використовуємо app.Evaluate, активувавши контекст потрібної книги
                try
                {
                    // Переконуємося, що Evaluate обчислюється для потрібної книги
                    string qualifiedExpression = $"'{wb.Name}'!{rangeExpression}";
                    object evaluated = app.Evaluate(qualifiedExpression);

                    if (evaluated is Range rng)
                    {
                        return rng;
                    }
                }
                catch
                {
                    // Резервний варіант: виклик через прямого власника
                    object evaluated = app.Evaluate(rangeExpression);
                    if (evaluated is Range rng)
                    {
                        return rng;
                    }
                }
            }

            return null;
        }

        #endregion

        #region Parsing: cell values (спільне для COM і ClosedXML)

        /// <summary>
        /// Збирає запис з значень комірок рядка (спільне для COM і ClosedXML). Номер колонки — 1-based.
        /// </summary>
        private static Item ParseItem(Func<int, object> cell, JobValue.Registry jobValue,
            PresenceValue.Registry presenceValue)
        {
            return new Item(ParseGuid(cell(ColPersonId)))
            {
                LineNumber = ParseInt(cell(1)),
                UnitCode = ParseInt(cell(2)),
                UnitName = ParseString(cell(3)),
                PositionId = ParseInt(cell(4)),
                PositionName = ParseString(cell(5)),
                PersonRank = ParseString(cell(7)),
                PersonFullName = ParseString(cell(8)),
                PersonTaxCode = ParseUlong(cell(9)),
                Presence = ParsePresenceValue(cell(11), presenceValue),
                Job = ParseJobValue(cell(ColJob), jobValue),
                JobDetails = ParseString(cell(ColJobDetails)),
                JobStartDate = ParseDateTimeOffset(cell(14)),
                JobStartTime = ParseTimeSpan(cell(15)),
                JobFinishDate = ParseDateTimeOffset(cell(16)),
                JobFinishTime = ParseTimeSpan(cell(17))
            };
        }

        private static bool IsCellEmpty(object cellValue)
        {
            if (cellValue == null || cellValue is ExcelEmpty || cellValue == DBNull.Value)
            {
                return true;
            }

            return string.IsNullOrWhiteSpace(cellValue.ToString());
        }

        private static string ParseString(object cellValue)
        {
            return IsCellEmpty(cellValue) ? string.Empty : cellValue.ToString().Trim();
        }

        private static int ParseInt(object cellValue)
        {
            if (IsCellEmpty(cellValue)) return 0;

            if (cellValue is double dVal) return (int)Math.Round(dVal);
            if (int.TryParse(cellValue.ToString(), out int result)) return result;

            return 0;
        }

        private static ulong ParseUlong(object cellValue)
        {
            if (IsCellEmpty(cellValue)) return 0;

            if (cellValue is double dVal) return (ulong)Math.Round(dVal);
            if (ulong.TryParse(cellValue.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out ulong result))
            {
                return result;
            }

            return 0;
        }

        private static Guid ParseGuid(object cellValue)
        {
            if (IsCellEmpty(cellValue)) return Guid.Empty;

            if (Guid.TryParse(cellValue.ToString(), out Guid result))
            {
                return result;
            }

            return Guid.Empty;
        }

        private static DateTimeOffset? ParseDateTimeOffset(object cellValue)
        {
            if (IsCellEmpty(cellValue)) return null;

            try
            {
                // Excel Value2 returns dates and times as OLE Automation numeric double values
                if (cellValue is double dVal)
                {
                    DateTime dt = DateTime.FromOADate(dVal);
                    return new DateTimeOffset(dt, TimeSpan.Zero);
                }

                // If cell was formatted as plain text
                if (DateTimeOffset.TryParse(cellValue.ToString(), CultureInfo.InvariantCulture, DateTimeStyles.None,
                        out DateTimeOffset parsedDt))
                {
                    return parsedDt;
                }
            }
            catch
            {
                // Return null if parsing fails
            }

            return null;
        }

        private static JobValue ParseJobValue(object cellValue, JobValue.Registry jobValue)
        {
            if (IsCellEmpty(cellValue)) return jobValue.Empty;

            string strVal = cellValue.ToString().Trim();
            if (string.IsNullOrWhiteSpace(strVal)) return jobValue.Empty;

            foreach (var jv in jobValue.AllValues)
            {
                if (strVal.Equals(jv.Value, StringComparison.InvariantCultureIgnoreCase))
                {
                    return jv;
                }
            }

            return jobValue.Unknown;
        }

        private static PresenceValue ParsePresenceValue(object cellValue, PresenceValue.Registry presenceValue)
        {
            if (IsCellEmpty(cellValue)) return presenceValue.Empty;

            string strVal = cellValue.ToString().Trim();
            if (string.IsNullOrWhiteSpace(strVal)) return presenceValue.Empty;

            foreach (var pv in presenceValue.AllValues)
            {
                if (strVal.Equals(pv.Value, StringComparison.InvariantCultureIgnoreCase))
                {
                    return pv;
                }
            }

            return presenceValue.Unknown;
        }

        private static TimeSpan? ParseTimeSpan(object cellValue)
        {
            if (IsCellEmpty(cellValue)) return null;

            try
            {
                // Case 1: Excel passed cell as double (fraction of a 24-hour day)
                if (cellValue is double dVal)
                {
                    DateTime dt = DateTime.FromOADate(dVal);
                    return dt.TimeOfDay;
                }

                string strVal = cellValue.ToString().Trim();

                // Case 2: Military string formats like "06:30", "6:30", "06:30:00"
                string[] timeFormats = new[] { @"hh\:mm", @"h\:mm", @"hh\:mm\:ss", @"h\:mm\:ss", "HH:mm", "H:mm" };
                if (TimeSpan.TryParseExact(strVal, timeFormats, CultureInfo.InvariantCulture, out TimeSpan parsedTime))
                {
                    return parsedTime;
                }

                // Case 3: Parse via DateTime exact match fallback
                if (DateTime.TryParseExact(strVal, timeFormats, CultureInfo.InvariantCulture, DateTimeStyles.None,
                        out DateTime dtResult))
                {
                    return dtResult.TimeOfDay;
                }

                // Case 4: General TimeSpan fallback
                if (TimeSpan.TryParse(strVal, CultureInfo.InvariantCulture, out TimeSpan fallbackTime))
                {
                    return fallbackTime;
                }
            }
            catch
            {
                // Return null if cell contains invalid time string format
            }

            return null;
        }

        #endregion
    }
}