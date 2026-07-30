#region

using System.Linq;
using NameCaseLib.NCL;

#endregion

namespace NameCaseLib.Core
{
    /// <summary>
    ///     Word - класс, который служит для хранения всей информации о каждом слове
    /// </summary>
    public class Word
    {
        /// <summary>
        /// Оригинальное слово
        /// </summary>
        public string WordOrig { get; }

        /// <summary>
        ///     Окончательное решение, к какому полу относится слово
        /// </summary>
        private Gender genderSolved = Gender.Null;

        /// <summary>
        ///     Содержит true, если все слово было в верхнем регистре и false, если не было
        /// </summary>
        private bool isUpperCase;

        /// <summary>
        ///     Маска больших букв в слове.
        ///     Содержит информацию о том, какие буквы в слове были большими, а какие маленькими:
        ///     - x - маленькая буква
        ///     - X - больная буква
        /// </summary>
        private LettersMask[] letterMask;

        /// <summary>
        ///     Создание нового объекта со словом
        /// </summary>
        /// <param name="word">Слово</param>
        public Word(string word)
        {
            WordOrig = word;
            GenerateMask(word);
            Name = word.ToLower();
        }

        /// <summary>
        ///     Массив содержит все падежи слова, полученные после склонения текущего слова
        /// </summary>
        public string[] NameCases { get; private set; }

        /// <summary>
        ///     Считывает или устанавливает все падежи
        /// </summary>
        public void SetNameCases(string[] nameCases, bool isReturnMask = true)
        {
            NameCases = nameCases;

            if (isReturnMask)
            {
                ReturnMask();
            }
        }

        /// <summary>
        ///     Рассчитывает и возвращает пол текущего слова. Или устанавливает нужный пол.
        /// </summary>
        public Gender Gender
        {
            get
            {
                if (genderSolved == Gender.Null)
                {
                    genderSolved = GenderProbability.Man >= GenderProbability.Woman ? Gender.Man : Gender.Woman;
                }

                return genderSolved;
            }
            set => genderSolved = value;
        }

        /// <summary>
        ///     Вероятность того, что текущей слово относится к или женскому полу
        /// </summary>
        public GenderProbability GenderProbability { get; set; }

        /// <summary>
        ///     Тип текущей записи (Фамилия/Имя/Отчество)
        /// </summary>
        public NamePart NamePart { get; set; } = NamePart.Null;

        /// <summary>
        ///     Слово в нижнем регистре, которое хранится в объекте класса
        /// </summary>
        public string Name { get; }

        /// <summary>
        ///     Номер правила, по которому было произведено склонение текущего слова
        /// </summary>
        public int Rule { get; set; }

        /// <summary>
        ///     Генерирует маску, которая содержит информацию о том, какие буквы в слове были большими, а какие маленькими:
        ///     - x - маленькая буква
        ///     - Х - большая буква
        /// </summary>
        /// <param name="word">Слово для которого нужна маска</param>
        private void GenerateMask(string word)
        {
            isUpperCase = true;
            var length = word.Length;
            letterMask = new LettersMask[length];

            for (var i = 0; i < length; i++)
            {
                var letter = word.Substring(i, 1);

                if (letter.All(char.IsLower))
                {
                    isUpperCase = false;
                    letterMask[i] = LettersMask.x;
                }
                else
                {
                    letterMask[i] = LettersMask.X;
                }
            }
        }

        /// <summary>
        ///     Возвращает все падежи слова в начальную маску
        /// </summary>
        private void ReturnMask()
        {
            var wordCount = NameCases.Length;

            if (isUpperCase)
            {
                for (var i = 0; i < wordCount; i++)
                {
                    NameCases[i] = NameCases[i].ToUpper();
                }
            }
            else
            {
                for (var i = 0; i < wordCount; i++)
                {
                    var lettersCount = NameCases[i].Length;
                    var maskLength = letterMask.Length;
                    var newStr = "";

                    for (var letter = 0; letter < lettersCount; letter++)
                    {
                        if (letter < maskLength && letterMask[letter] == LettersMask.X)
                        {
                            newStr += NameCases[i].Substring(letter, 1).ToUpper();
                        }
                        else
                        {
                            newStr += NameCases[i].Substring(letter, 1).ToLower();
                        }
                    }

                    NameCases[i] = newStr;
                }
            }
        }

        /// <summary>
        ///     Возвращает строку с нужным падежом текущего слова
        /// </summary>
        /// <param name="padeg">нужный падеж</param>
        /// <returns>строка с нужным падежом текущего слова</returns>
        public string GetNameCase(Padeg padeg)
        {
            return NameCases[(int)padeg];
        }

        /// <summary>
        ///     Если уже был рассчитан пол для всех слов системы, тогда каждому слову предается окончательное
        ///     решение. Эта функция определяет было ли принято окончательное решение.
        /// </summary>
        /// <returns>true если определен и false если нет</returns>
        public bool IsGenderSolved()
        {
            return genderSolved != Gender.Null;
        }
    }
}