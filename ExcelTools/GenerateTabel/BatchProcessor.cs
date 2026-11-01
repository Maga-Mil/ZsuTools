using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ZsuTools.GenerateTabel;

namespace ZsuTools
{
    public class BatchProcessor
    {
        public static Dictionary<int, List<Tuple<RankPerson, SZState, Color>>> ProcessMonthlyFiles()
        {
            // 1. Показуємо користувачу діалог вибору файлу (.NET Framework Windows Forms)
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Файли Excel (*.xlsx)|*.xlsx|Усі файли (*.*)|*.*";
                openFileDialog.Title = "Оберіть будь-який файл відомості для обробки місяця";

                if (openFileDialog.ShowDialog() != DialogResult.OK)
                {
                    return new Dictionary<int, List<Tuple<RankPerson, SZState, Color>>>(); // Користувач скасував вибір
                }

                string selectedFilePath = openFileDialog.FileName;
                string directory = Path.GetDirectoryName(selectedFilePath);
                string fileName = Path.GetFileName(selectedFilePath);

                // 2. Розбираємо ім'я файлу за допомогою Regex
                // Шаблон шукає: "16. РБАК " -> потім дві цифри дня -> крапка -> дві цифри місяця -> крапка -> чотири цифри року -> .xlsx
                // Використовуємо групи (скобки), щоб зафіксувати префікс, місяць та рік
                var regex = new Regex(
                    @"^(?<prefix>.+?)(?<day>\d{2})\.(?<month>\d{2})\.(?<year>\d{4})\.xlsx$",
                    RegexOptions.IgnoreCase);
                var match = regex.Match(fileName);

                if (!match.Success)
                {
                    MessageBox.Show("Назва вибраного файлу не відповідає формату '<префікс> ДД.ММ.РРРР.xlsx'. Неможливо",
                        "Помилка формату", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return new Dictionary<int, List<Tuple<RankPerson, SZState, Color>>>();
                }

                string prefix = match.Groups["prefix"].Value;
                string monthStr = match.Groups["month"].Value;
                string yearStr = match.Groups["year"].Value;

                // Сюди збираємо результати: Ключ — день місяця, Значення — список tuples людей
                var monthlyResults = new Dictionary<int, List<Tuple<RankPerson, SZState, Color>>>();

                try
                {
                    // 3. Цикл по всіх можливих днях місяця (від 1 до 31)
                    for (int day = 1; day <= 31; day++)
                    {
                        // Форматуємо день як "01", "02" і т.д.
                        string dayStr = day.ToString("D2");
                        string targetFileName = string.Format("{0}{1}.{2}.{3}.xlsx", prefix, dayStr, monthStr, yearStr);
                        string targetFilePath = Path.Combine(directory, targetFileName);

                        // Якщо місяць неповний або дня немає — просто пропускаємо його
                        if (!File.Exists(targetFilePath))
                        {
                            continue;
                        }

                        try
                        {
                            // ClosedXML читає файл напряму, без Excel Interop
                            var stroyova = new СтройоваЗаписка(targetFilePath);
                            monthlyResults.Add(day, stroyova.MilitaryData);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Помилка обробки файлу {targetFileName}. : {ex.Message}", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }

                    // 4. Обробка фінального результату
                    //ProcessFinalResults(monthlyResults);
                    return monthlyResults;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Помилка пакетної обробки: " + ex.Message);
                }
            }

            return  new Dictionary<int, List<Tuple<RankPerson, SZState, Color>>>();
        }
        
    }
}