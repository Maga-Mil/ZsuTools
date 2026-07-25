using System;
using System.Text.RegularExpressions;

namespace ZsuTools
{
    public class RankPerson
    {
        // Статичний скомпільований Regex для кращої швидкodії (Thread-Safe)
        private static readonly Regex NameParserRegex = new Regex(
            @"(?<fullName>(?:[А-ЯЇЄІҐ][а-яїєіґʼ’'\w-]+\s+){2}[А-ЯЇЄІҐ][а-яїєіґʼ’'\w-]+)$",
            RegexOptions.Compiled);

        public string Rank { get; }
        public string FullName { get; }

        /// <summary>
        /// Конструктор, який автоматично розбирає вхідний рядок на Звання та ПІБ
        /// </summary>
        public RankPerson(string rawInput)
        {
            if (string.IsNullOrEmpty(rawInput))
            {
                Rank = string.Empty;
                FullName = string.Empty;
                return;
            }

            // 1. Очищаємо від подвійних пробілів
            string cleanedInput = Regex.Replace(rawInput.Trim(), @"\s+", " ");

            // 2. Спроба розпарсити через Regex
            var match = NameParserRegex.Match(cleanedInput);

            if (match.Success)
            {
                FullName = match.Groups["fullName"].Value.Trim();
                Rank = cleanedInput.Substring(0, match.Index).Trim();
            }
            else
            {
                // Резервний варіант (Fallback), якщо ПІБ не відповідає стандарту 3 слів
                string[] words = cleanedInput.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                if (words.Length <= 3)
                {
                    Rank = string.Empty;
                    FullName = string.Join(" ", words);
                }
                else
                {
                    FullName = string.Join(" ", words, words.Length - 3, 3);
                    Rank = string.Join(" ", words, 0, words.Length - 3);
                }
            }
        }

        // --- Перевизначення для безпечного використання в Dictionary ---

        public override bool Equals(object obj)
        {
            if (obj is RankPerson other)
            {
                // Порівнюємо людей за ПІБ без урахування регістру
                return string.Equals(this.FullName, other.FullName, StringComparison.InvariantCultureIgnoreCase);
            }

            return false;
        }

        public override int GetHashCode()
        {
            // Hash-код беремо від FullName в нижньому регістрі, щоб він збігався для "Іванов" та "ІВАНОВ"
            return FullName != null ? StringComparer.InvariantCultureIgnoreCase.GetHashCode(FullName) : 0;
        }

        public override string ToString()
        {
            return string.IsNullOrEmpty(Rank) ? FullName : $"{Rank} {FullName}";
        }
    }
}