using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;

namespace ZsuTools.Tables
{
    public class TabelZaluchenosti
    {
        public List<Item> Records { get; private set; } = new List<Item>();

        public JobValue.Registry Jobs { get; private set; }
        public PresenceValue.Registry Presence { get; private set; }

        public TabelZaluchenosti(Workbook wb)
        {
            var jobsValues = GetDictionaryFromNamedRanges(wb, "InvVal[val_code]", "InvVal[val_name]");
            Jobs = new JobValue.Registry(jobsValues);

            var presenceValues = GetDictionaryFromNamedRanges(wb, "PersonVal[val_code]", "PersonVal[val_name]");
            Presence = new PresenceValue.Registry(presenceValues);

            var mainWorkSheet = ExcelUtils.FindWorksheetByNameContains(wb, "Табелювання");
            var items = ReadTableFromWorksheet(mainWorkSheet, Jobs, Presence);
            Records.AddRange(items);
        }

        /// <summary>
        /// Reads table data directly from the provided <see cref="Worksheet"/> instance parameter.
        /// </summary>
        /// <param name="sheet">The target Excel worksheet instance to read from.</param>
        /// <returns>List of parsed <see cref="Item"/> items.</returns>
        private static List<Item> ReadTableFromWorksheet(Worksheet sheet, JobValue.Registry jobValue, PresenceValue.Registry presenceValue)
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

                var record = new Item
                {
                    LineNumber = ParseInt(data[row, 1]),
                    UnitCode = ParseInt(data[row, 2]),
                    UnitName = ParseString(data[row, 3]),
                    PositionId = ParseInt(data[row, 4]),
                    PositionName = ParseString(data[row, 5]),
                    PersonId = ParseGuid(data[row, 6]),
                    PersonRank = ParseString(data[row, 7]),
                    PersonFullName = ParseString(data[row, 8]),
                    PersonTaxCode = ParseUlong(data[row, 9]),
                    Presence = ParsePresenceValue(data[row, 11], presenceValue),
                    Job = ParseJobValue(data[row, 12], jobValue),
                    JobDetails = ParseString(data[row, 13]),
                    JobStartDate = ParseDateTimeOffset(data[row, 14]),
                    JobStartTime = ParseTimeSpan(data[row, 15]),
                    JobFinishDate = ParseDateTimeOffset(data[row, 16]),
                    JobFinishTime = ParseTimeSpan(data[row, 17])
                };

                results.Add(record);
            }

            return results;
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

        public static List<string> GetNamedRangeValues(Workbook workbook, string namedRangeName)
        {
            var list = new List<string>();
            if (workbook == null || string.IsNullOrWhiteSpace(namedRangeName)) return list;

            try
            {
                Name namedRange = workbook.Names.Item(namedRangeName);
                Range range = namedRange?.RefersToRange;
                if (range != null)
                {
                    object[,] values = range.Value2 as object[,];
                    if (values != null)
                    {
                        int rows = values.GetLength(0);
                        int cols = values.GetLength(1);
                        for (int r = 1; r <= rows; r++)
                        {
                            for (int c = 1; c <= cols; c++)
                            {
                                string val = values[r, c]?.ToString()?.Trim();
                                if (!string.IsNullOrEmpty(val))
                                {
                                    list.Add(val);
                                }
                            }
                        }
                    }
                    else if (range.Value2 != null)
                    {
                        list.Add(range.Value2.ToString().Trim());
                    }
                }
            }
            catch
            {
                // Named range not found or invalid formula reference
            }

            return list;
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

        public class Item
        {
            public int LineNumber { get; set; }
            public int UnitCode { get; set; }
            public string UnitName { get; set; }
            public int PositionId { get; set; }
            public string PositionName { get; set; }
            public Guid PersonId { get; set; }
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
                    TRIP = Create(values, nameof(TRIP));
                    AWOL = Create(values, nameof(AWOL));

                    AllValues = new[]
                    {
                        BATTALION_DEFENSE_SECTOR,
                        COMBAT_SUPPORT,
                        VACATION,
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
    }
}