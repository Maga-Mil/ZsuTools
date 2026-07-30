using System;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using ExcelDna.Integration;
using NameCaseLib;
using NameCaseLib.NCL;

namespace ZsuTools
{
    public class StringFunctions
    {
        public enum EWordForm
        {
            Possessive = 0,
            Dative = 1
        }
        
        [ExcelFunction(
            Name = "ToPossesiveCase", // Назва функції в Excel (опціонально, за замовчуванням дорівнює імені методу)
            Description = "Перетворює звання/посаду/підрозділ у родовий відмінок",
            Category = "ZsuTools" // Функція буде згрупована у цій категорії в Excel
            )]
        public static object ToPossesiveCase(
            [ExcelArgument(Name = "звання/посада/підрозділ", Description = "Текст для перетворення в родовий відмінок")]
            string input)
        {
            // Apply all configured replacements (includes взвод/рота in the map)
            return ApplyReplacements(input);
        }
        
        [ExcelFunction(
            Name = "NameToPossesiveCase", // Назва функції в Excel (опціонально, за замовчуванням дорівнює імені методу)
            Description = "Перетворює ПІБ у родовий відмінок",
            Category = "ZsuTools" // Функція буде згрупована у цій категорії в Excel
        )]
        public static object NameToPossesiveCase(
            [ExcelArgument(Name = "ПІБ", Description = "Текст для перетворення в родовий відмінок")]
            string input)
        {
            var uaFormer = new Ua();
            return uaFormer.QFullName(input, Padeg.UaRodovyi);
        }
        
        [ExcelFunction(
            Name = "ToDativeCase", // Назва функції в Excel (опціонально, за замовчуванням дорівнює імені методу)
            Description = "Перетворює звання/посаду/підрозділ у давальний відмінок",
            Category = "ZsuTools" // Функція буде згрупована у цій категорії в Excel
        )]
        public static object ToDativeCase(
            [ExcelArgument(Name = "підрозділ", Description = "Текст для перетворення в давальний відмінок")]
            string input)
        {
            // Apply all configured replacements (includes взвод/рота in the map)
            return ApplyReplacements(input, EWordForm.Dative);
        }
        
        // Mapping of source words to their forms: (possessive/genitive, dative). Keys are in lowercase.
        private static readonly Dictionary<string, (string Possessive, string Dative)> WordForms = new Dictionary<string, (string Possessive, string Dative)>
        {
            { "взвод", ("взводу", "взводу") },
            { "рота",  ("роти",  "роті") },
            { "солдат", ("солдата", "солдату") },
            { "молодший", ("молодшого", "молодшому") },
            { "сержант", ("сержанта", "сержанту") },
            { "старший", ("старшого", "старшому") },
            { "лейтенант", ("лейтенанта", "лейтенанту") },
            { "головний", ("головного", "головному") },
            { "оператор", ("оператора", "оператору") },
            { "електрик", ("електрика", "електрику") },
            { "водій", ("водія", "водію") },
            { "командир", ("командира", "командиру") },
            { "механік", ("механіка", "механіку") },
            { "майстер", ("майстра", "майстру") },
            { "дешифрувальник", ("дешифрувальника", "дешифрувальнику") },
            { "бойовий", ("бойового", "бойовому") },
            { "медик", ("медика", "медику") },
            { "технік", ("техніка", "техніку") },
        };

        // Precompiled regexes for each word to find (initialized once in static ctor)
        private static readonly Dictionary<string, Regex> WordFindRegexes;

        static StringFunctions()
        {
            WordFindRegexes = new Dictionary<string, Regex>(StringComparer.OrdinalIgnoreCase);
            foreach (var key in WordForms.Keys)
            {
                var rx = new Regex($"\\b{Regex.Escape(key)}\\b", RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);
                WordFindRegexes[key] = rx;
            }
        }

        // Apply replacements while preserving case style:
        // - ALL CAPS -> ALL CAPS replacement
        // - First letter uppercase -> Capitalized replacement
        // - otherwise -> lowercase replacement
        private static string ApplyReplacements(string input, EWordForm form = EWordForm.Possessive)
        {
            if (string.IsNullOrEmpty(input)) return input;

            string result = input;
            foreach (var kv in WordForms)
            {
                string word = kv.Key; // lowercase key
                var forms = kv.Value; // (Possessive, Dative)

                // use precompiled regex
                if (!WordFindRegexes.TryGetValue(word, out var regex))
                    continue;

                // choose form
                string replacementLower = form == EWordForm.Dative ? forms.Dative : forms.Possessive;

                result = regex.Replace(result, match =>
                {
                    string val = match.Value;
                    string replacementUpper = replacementLower.ToUpperInvariant();

                    if (IsAllUpperInvariant(val)) // all uppercase
                        return replacementUpper;

                    if (IsCapitalizedInvariant(val)) // first letter uppercase, rest lowercase
                        return CapitalizeInvariant(replacementLower);

                    return replacementLower;
                });
            }

            return result;
        }

        private static string CapitalizeInvariant(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            if (s.Length == 1) return s.ToUpperInvariant();
            return char.ToUpperInvariant(s[0]) + s.Substring(1).ToLowerInvariant();
        }

        // Returns true if the string contains at least one letter and every letter is uppercase (invariant)
        private static bool IsAllUpperInvariant(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            bool hasLetter = false;
            foreach (char c in s)
            {
                if (char.IsLetter(c))
                {
                    hasLetter = true;
                    if (char.ToUpperInvariant(c) != c) return false;
                }
            }
            return hasLetter;
        }

        // Returns true if the first letter is uppercase and all other letters (if any) are lowercase
        private static bool IsCapitalizedInvariant(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            int firstLetterIndex = -1;
            for (int i = 0; i < s.Length; i++)
            {
                if (char.IsLetter(s[i]))
                {
                    firstLetterIndex = i;
                    break;
                }
            }
            if (firstLetterIndex == -1) return false;
            // first letter must be uppercase
            if (char.ToUpperInvariant(s[firstLetterIndex]) != s[firstLetterIndex]) return false;
            // remaining letters (letters only) must be lowercase
            for (int i = firstLetterIndex + 1; i < s.Length; i++)
            {
                char c = s[i];
                if (char.IsLetter(c) && char.ToLowerInvariant(c) != c) return false;
            }
            return true;
        }

    }
}