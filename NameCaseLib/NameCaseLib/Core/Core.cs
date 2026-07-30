#region

using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using NameCaseLib.NCL;

#endregion

namespace NameCaseLib.Core
{
    /// <summary>
    ///     Набор основных функций, который позволяют сделать интерфейс склонения русского и украинского языка
    ///     абсолютно одинаковым. Содержит все функции для внешнего взаимодействия с библиотекой.
    /// </summary>
    public abstract class Core
    {
        /// <summary>
        ///     Версия языкового файла
        /// </summary>
        protected static string LanguageBuild;

        /// <summary>
        ///     Количество падежей в языке
        /// </summary>
        protected abstract int CaseCount { get; }

        /// <summary>
        ///     Если все текущие слова были просклонены и в каждом слове уже есть результат склонения,
        ///     тогда true. Если было добавлено новое слово флаг сбрасывается на false
        /// </summary>
        private bool isFinished;

        /// <summary>
        ///     Массив содержит результат склонения слова - слово во всех падежах
        /// </summary>
        protected string[] LastResult;

        /// <summary>
        ///     Номер последнего использованного правила, устанавливается методом Rule()
        /// </summary>
        protected int LastRule { get; set; }

        /// <summary>
        ///     Готовность системы:
        ///     - Все слова идентифицированы (известно к какой части ФИО относится слово)
        ///     - У всех слов определен пол
        ///     Если все сделано стоит флаг true, при добавлении нового слова флаг сбрасывается на false
        /// </summary>
        private bool isReady;

        /// <summary>
        ///     Метод Last() вырезает подстроки разной длины. Поскольку одинаковых вызовов бывает несколько,
        ///     то все результаты выполнения кешируются в этом массиве.
        /// </summary>
        private Dictionary<Tuple<int, int>, string> workingLastCache;

        /// <summary>
        ///     Переменная, в которую заносится слово с которым сейчас идет работа
        /// </summary>
        protected string WorkingWord;

        /// <summary>
        ///     Возвращает текущую версию библиотеки
        /// </summary>
        public const string Version = "0.2/0.4.1";

        /// <summary>
        ///     Возвращает текущую версию языкового файла
        /// </summary>
        public abstract string LanguageVersion { get; }

        /// <summary>
        ///     Метод очищает результаты последнего склонения слова. Нужен при склонении нескольких слов.
        /// </summary>
        private void Reset()
        {
            LastRule = 0;
            LastResult = new string[CaseCount];
        }

        /// <summary>
        ///     Устанавливает флаги о том, что система не готово и слова еще не были просклонены
        /// </summary>
        private void NotReady()
        {
            isReady = false;
            isFinished = false;
        }

        /// <summary>
        ///     Сбрасывает все информацию на начальную. Очищает все слова добавленные в систему.
        ///     После выполнения система готова работать с начала.
        /// </summary>
        public Core FullReset()
        {
            Words = new List<Word>();
            Reset();
            NotReady();

            return this;
        }

        /// <summary>
        ///     Устанавливает слово текущим для работы системы. Очищает кеш слова.
        /// </summary>
        /// <param name="word">слово, которое нужно установить</param>
        protected void SetWorkingWord(string word)
        {
            Reset();
            WorkingWord = word;
            workingLastCache = new Dictionary<Tuple<int, int>, string>();
        }

        /// <summary>
        ///     Если не нужно склонять слово, делает результат таким же как и именительный падеж
        /// </summary>
        protected void MakeResultTheSame()
        {
            LastResult = new string[CaseCount];

            for (var i = 0; i < CaseCount; i++)
            {
                LastResult[i] = WorkingWord;
            }
        }

        /// <summary>
        ///     Вырезает определенное количество последних букв текущего слова
        /// </summary>
        /// <param name="length">Количество букв</param>
        /// <returns>Подстроку содержащую определенное количество букв</returns>
        protected string Last(int length)
        {
            var key = new Tuple<int, int>(length, length);

            if (!workingLastCache.TryGetValue(key, out var result))
            {
                var startIndex = WorkingWord.Length - length;

                result = startIndex >= 0 ? WorkingWord.Substring(WorkingWord.Length - length, length) : WorkingWord;

                workingLastCache.Add(key, result);
            }

            return result;
        }

        /// <summary>
        ///     Получает указанное количество букв с конца слова
        /// </summary>
        /// <param name="word">Слово</param>
        /// <param name="length">Количество букв</param>
        /// <returns>Полученная строка</returns>
        protected static string Last(string word, int length)
        {
            var startIndex = word.Length - length;

            var result = startIndex >= 0 ? word.Substring(word.Length - length, length) : word;

            return result;
        }

        /// <summary>
        ///     Вырезает stopAfter букв начиная с length с конца
        /// </summary>
        /// <param name="length">На сколько букв нужно отступить от конца</param>
        /// <param name="stopAfter">Сколько букв нужно вырезать</param>
        /// <returns>Искомая строка</returns>
        protected string Last(int length, int stopAfter)
        {
            var key = new Tuple<int, int>(length, stopAfter);

            if (!workingLastCache.TryGetValue(key, out var result))
            {
                var startIndex = WorkingWord.Length - length;

                result = startIndex >= 0 ? WorkingWord.Substring(WorkingWord.Length - length, stopAfter) : WorkingWord;

                workingLastCache.Add(key, result);
            }

            return result;
        }

        /// <summary>
        ///     Извлекает последние буквы из указанного слова
        /// </summary>
        /// <param name="word">Слово</param>
        /// <param name="length">Сколько букв с конца</param>
        /// <param name="stopAfter">Сколько нужно взять</param>
        /// <returns>Полученная подстрока</returns>
        protected static string Last(string word, int length, int stopAfter)
        {
            var startIndex = word.Length - length;

            var result = startIndex >= 0 ? word.Substring(word.Length - length, stopAfter) : word;

            return result;
        }

        /// <summary>
        ///     Над текущим словом выполняются правила в указанном порядке.
        /// </summary>
        /// <param name="gender">Пол текущего слова</param>
        /// <param name="rulesArray">Порядок правил</param>
        /// <returns>Если правило было использовано true если нет тогда false</returns>
        protected bool RulesChain(Gender gender, int[] rulesArray)
        {
            if (gender != Gender.Null)
            {
                var rulesLength = rulesArray.Length;
                var rulesName = gender == Gender.Man ? "Man" : "Woman";
                var classType = GetType();

                for (var i = 0; i < rulesLength; i++)
                {
                    var methodName = $"{rulesName}Rule{rulesArray[i]}";

                    var res = (bool)classType.GetMethod(methodName, BindingFlags.NonPublic | BindingFlags.Instance)
                        .Invoke(this, null);

                    if (res)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        ///     Проверяет входит ли буква в список букв
        /// </summary>
        /// <param name="letter">буква</param>
        /// <param name="letters">список букв</param>
        /// <returns>true если входит в список и false если не входит</returns>
        protected static bool InLetters(string letter, string letters)
        {
            if (letter != "")
            {
                if (letters.IndexOf(letter, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        ///     Ищет окончание в списке окончаний
        /// </summary>
        /// <param name="ending">окончание</param>
        /// <param name="endings">список окончаний</param>
        /// <returns>true если найдено и false если нет</returns>
        protected static bool InEndings(string ending, params string[] endings)
        {
            if (ending != "")
            {
                return endings.Contains(ending, StringComparer.OrdinalIgnoreCase);
            }

            return false;
        }

        /// <summary>
        ///     Проверяет входит ли имя в список имен
        /// </summary>
        /// <param name="name">имя</param>
        /// <param name="names">список имен</param>
        /// <returns>true если входит</returns>
        protected static bool InNames(string name, params string[] names)
        {
            return names.Contains(name, StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>
        ///     Склоняем слово во все падежи используя окончания
        /// </summary>
        /// <param name="word">Слово</param>
        /// <param name="replaceLast">сколько последних букв нужно убрать</param>
        /// <param name="endings">окончания</param>
        protected void WordForms(string word, int replaceLast, params string[] endings)
        {
            //Сохраняем именительный падеж
            LastResult = new string[CaseCount];
            LastResult[0] = WorkingWord;

            word = word.Length >= replaceLast ? word.Substring(0, word.Length - replaceLast) : "";

            // Приписываем окончания
            for (var i = 1; i < CaseCount; i++)
            {
                LastResult[i] = word + endings[i - 1];
            }
        }

        /// <summary>
        ///     Создает список слов во всех падежах используя окончания для каждого падежа
        /// </summary>
        /// <param name="word">слово</param>
        /// <param name="endings">окончания</param>
        protected void WordForms(string word, params string[] endings)
        {
            WordForms(word, 0, endings);
        }

        /// <summary>
        ///     Установить имя человека
        /// </summary>
        /// <param name="name">Имя</param>
        /// <returns></returns>
        public Core SetName(string name)
        {
            if (name.Trim() != "")
            {
                var tmpWord = new Word(name)
                {
                    NamePart = NamePart.Name
                };

                Words.Add(tmpWord);
                NotReady();
            }

            return this;
        }

        /// <summary>
        ///     Установить фамилию человека
        /// </summary>
        /// <param name="surname">Фамилия</param>
        /// <returns></returns>
        public Core SetSurname(string surname)
        {
            if (surname.Trim() != "")
            {
                var tmpWord = new Word(surname)
                {
                    NamePart = NamePart.Surname
                };

                Words.Add(tmpWord);
                NotReady();
            }

            return this;
        }

        /// <summary>
        ///     Установить отчество человека
        /// </summary>
        /// <param name="fatherName">Отчество</param>
        /// <returns></returns>
        public Core SetFatherName(string fatherName)
        {
            if (fatherName.Trim() != "")
            {
                var tmpWord = new Word(fatherName)
                {
                    NamePart = NamePart.FatherName
                };

                Words.Add(tmpWord);
                NotReady();
            }

            return this;
        }

        /// <summary>
        ///     Устанавливает пол человека
        /// </summary>
        /// <param name="gender">Пол человека</param>
        /// <returns></returns>
        public Core SetGender(Gender gender)
        {
            foreach (var w in Words)
            {
                w.Gender = gender;
            }

            return this;
        }

        /// <summary>
        ///     Устанавливает полное ФИО
        /// </summary>
        /// <param name="surname">Фамилия</param>
        /// <param name="name">Имя</param>
        /// <param name="fatherName">Отчество</param>
        /// <returns></returns>
        public Core SetFullName(string surname, string name, string fatherName)
        {
            SetName(name);
            SetSurname(surname);
            SetFatherName(fatherName);

            return this;
        }

        /// <summary>
        ///     Идентифицирует нужное слово
        /// </summary>
        /// <param name="word">Слово</param>
        private void PrepareNamePart(Word word)
        {
            if (word.NamePart == NamePart.Null)
            {
                DetectNamePart(word);
            }
        }

        /// <summary>
        ///     Идентифицирует все существующие слова
        /// </summary>
        private void PrepareAllNameParts()
        {
            foreach (var w in Words)
            {
                PrepareNamePart(w);
            }
        }

        /// <summary>
        ///     Предварительно определяет пол во слове
        /// </summary>
        /// <param name="word">Слово для определения</param>
        private void PrepareGender(Word word)
        {
            if (!word.IsGenderSolved())
            {
                switch (word.NamePart)
                {
                    case NamePart.Name:
                        GenderByName(word);

                        break;
                    case NamePart.Surname:
                        GenderBySurname(word);

                        break;
                    case NamePart.FatherName:
                        GenderByFatherName(word);

                        break;
                }
            }
        }

        /// <summary>
        ///     Принимает решение о текущем поле человека
        /// </summary>
        private void SolveGender()
        {
            //Ищем, может где-то пол уже установлен

            foreach (var w in Words.Where(w => w.IsGenderSolved()))
            {
                SetGender(w.Gender);

                return;
            }

            //Если нет тогда определяем у каждого слова и потом суммируем
            var probability = new GenderProbability(0, 0);

            foreach (var word in Words)
            {
                PrepareGender(word);
                probability += word.GenderProbability;
            }

            SetGender(probability.Man > probability.Woman ? Gender.Man : Gender.Woman);
        }

        /// <summary>
        ///     Выполняет все необходимые подготовления для склонения.
        ///     Все слова идентифицируются. Определяется пол.
        /// </summary>
        private void PrepareEverything()
        {
            if (!isReady)
            {
                PrepareAllNameParts();
                SolveGender();
                isReady = true;
            }
        }

        /// <summary>
        ///     По указанным словам определяется пол человека
        /// </summary>
        /// <returns>Пол человека</returns>
        private Gender GenderAutoDetect()
        {
            PrepareEverything();

            var probability = new GenderProbability(0, 0);

            foreach (var w in Words)
            {
                probability += w.GenderProbability;
            }

            if (probability.Man > probability.Woman)
            {
                return Gender.Man;
            }

            if (probability.Woman > probability.Man)
            {
                return Gender.Woman;
            }

            return Gender.Null;
        }

        /// <summary>
        ///     Разделяет слова на части и готовит к дальнейшему склонению
        /// </summary>
        /// <param name="fullname">Строка которая содержит полное имя</param>
        private void SplitFullName(string fullname)
        {
            var arr = fullname.Trim().Split(' ');
            var length = arr.Length;

            Words = new List<Word>();

            for (var i = 0; i < length; i++)
            {
                if (!string.IsNullOrEmpty(arr[i]))
                {
                    Words.Add(new Word(arr[i]));
                }
            }
        }

        /// <summary>
        ///     Разбивает строку на слова и возвращает формат в котором записано имя
        ///     <br/>
        ///     <b>Формат:</b>
        ///     - S - Фамилия
        ///     - N - Имя
        ///     - F - Отчество
        /// </summary>
        /// <param name="fullName">Строка, для которой необходимо определить формат</param>
        /// <returns>Формат в котором записано имя</returns>
        public string GetFullNameFormat(string fullName)
        {
            FullReset();
            SplitFullName(fullName);

            return Words.Aggregate(string.Empty, (current, w) => current + $"{w.NamePart}");
        }

        /// <summary>
        ///     Склоняет слово по нужным правилам
        /// </summary>
        /// <param name="word">Слово</param>
        protected virtual void WordCase(Word word)
        {
            var method = GetType()
                .GetMethod($"{word.Gender.ToString("g")}{word.NamePart.ToString("g")}",
                    BindingFlags.NonPublic | BindingFlags.Instance);

            //если фамилия из 2х слов через дефис
            //http://new.gramota.ru/spravka/buro/search-answer?s=273912
            //разбиваем слово с дефисами на части

            if (word.NamePart == NamePart.Surname && word.WordOrig.Contains('-'))
            {
                var result = new string[CaseCount];
                var lastRule = -1;
                var currentWords = word.WordOrig.Split('-');
                var oCurWords = new List<Word>(currentWords.Length);

                for (var k = 0; k < currentWords.Length; k++)
                {
                    var curWord = currentWords[k];
                    var isNormRules = true;
                    var oNcw = new Word(curWord);

                    if (currentWords.Length > 1 && k < currentWords.Length - 1)
                    {
                        //если первая часть фамилии тоже фамилия, то склоняем по общим правилам
                        //иначе не склоняется

                        if (!InNames(curWord, "Тулуз"))
                        {
                            var instance = (Core)Activator.CreateInstance(GetType());
                            instance.DetectNamePart(oNcw);
                            isNormRules = oNcw.NamePart == word.NamePart;
                        }
                        else
                        {
                            isNormRules = false;
                        }
                    }

                    SetWorkingWord(curWord);

                    string[] tmpResult;

                    if (isNormRules && (bool)method.Invoke(this, null))
                    {
                        //склоняется
                        tmpResult = LastResult;
                        lastRule = LastRule;
                    }
                    else
                    {
                        //не склоняется. Заполняем что есть
                        tmpResult = Enumerable.Repeat(curWord, CaseCount).ToArray();
                        lastRule = -1;
                    }

                    oNcw.SetNameCases(tmpResult);
                    oCurWords.Add(oNcw);
                }

                foreach (var nameCases in oCurWords.Select(oNcw => oNcw.NameCases))
                {
                    for (var k = 0; k < nameCases.Length; k++)
                    {
                        var nameCase = nameCases[k];

                        if (!string.IsNullOrEmpty(result[k]))
                        {
                            result[k] = $"{result[k]}-{nameCase}";
                        }
                        else
                        {
                            result[k] = nameCase;
                        }
                    }
                }

                //устанавливаем падежи для целого слова
                word.SetNameCases(result, false);

                word.Rule = lastRule;
            }
            else
            {
                SetWorkingWord(word.Name);

                if ((bool)method.Invoke(this, null))
                {
                    word.SetNameCases(LastResult);
                    word.Rule = LastRule;
                }
                else
                {
                    MakeResultTheSame();
                    word.SetNameCases(LastResult);
                    word.Rule = -1;
                }
            }
        }

        /// <summary>
        ///     Производит склонение всех слов
        /// </summary>
        private void AllWordCases()
        {
            if (!isFinished)
            {
                PrepareEverything();

                foreach (var w in Words)
                {
                    WordCase(w);
                }

                isFinished = true;
            }
        }

        /// <summary>
        ///     Возвращает массив который содержит все падежи имени
        /// </summary>
        /// <returns>Возвращает массив со всеми падежами имени</returns>
        public string[] GetNameCases()
        {
            AllWordCases();

            return Words.FirstOrDefault(w => w.NamePart == NamePart.Name)?.NameCases ?? new string[CaseCount];
        }

        /// <summary>
        ///     Возвращает имя в определенном падеже
        /// </summary>
        /// <param name="caseNum">Падеж</param>
        /// <returns>Имя в определенном падеже</returns>
        public string GetNameCase(Padeg caseNum)
        {
            AllWordCases();

            return Words.FirstOrDefault(w => w.NamePart == NamePart.Name)?.GetNameCase(caseNum) ?? string.Empty;
        }

        /// <summary>
        ///     Возвращает массив который содержит все падежи фамилии
        /// </summary>
        /// <returns>Возвращает массив со всеми падежами фамилии</returns>
        public string[] GetSurnameCases()
        {
            AllWordCases();

            return Words.FirstOrDefault(w => w.NamePart == NamePart.Surname)?.NameCases ?? new string[CaseCount];
        }

        /// <summary>
        ///     Возвращает фамилию в определенном падеже
        /// </summary>
        /// <param name="caseNum">Падеж</param>
        /// <returns>Фамилия в определенном падеже</returns>
        public string GetSurnameCase(Padeg caseNum)
        {
            AllWordCases();

            return Words.FirstOrDefault(w => w.NamePart == NamePart.Surname)?.GetNameCase(caseNum) ?? string.Empty;
        }

        /// <summary>
        ///     Возвращает массив который содержит все падежи отчества
        /// </summary>
        /// <returns>Возвращает массив со всеми падежами отчества</returns>
        public string[] GetFatherNameCases()
        {
            AllWordCases();

            return Words.FirstOrDefault(w => w.NamePart == NamePart.FatherName)?.NameCases ?? new string[CaseCount];
        }

        /// <summary>
        ///     Возвращает отчество в определенном падеже
        /// </summary>
        /// <param name="caseNum">Падеж</param>
        /// <returns>Отчество в определенном падеже</returns>
        public string GetFatherNameCase(Padeg caseNum)
        {
            AllWordCases();

            return Words.FirstOrDefault(w => w.NamePart == NamePart.FatherName)?.GetNameCase(caseNum) ?? string.Empty;
        }

        /// <summary>
        ///     Выполняет склонение имени
        /// </summary>
        /// <param name="name">Имя</param>
        /// <param name="gender">Пол</param>
        /// <returns>Массив со всеми падежами</returns>
        public string[] QName(string name, Gender gender)
        {
            FullReset();
            SetName(name);

            if (gender != Gender.Null)
            {
                SetGender(gender);
            }

            return GetNameCases();
        }

        /// <summary>
        ///     Выполняет склонение имени
        /// </summary>
        /// <param name="name">Имя</param>
        /// <returns>Массив со всеми падежами</returns>
        public string[] QName(string name)
        {
            return QName(name, Gender.Null);
        }

        /// <summary>
        ///     Выполняет склонение имени
        /// </summary>
        /// <param name="name">Имя</param>
        /// <param name="caseNum">Падеж</param>
        /// <param name="gender">Пол</param>
        /// <returns>Имя в указанном падеже</returns>
        public string QName(string name, Padeg caseNum, Gender gender)
        {
            FullReset();
            SetName(name);

            if (gender != Gender.Null)
            {
                SetGender(gender);
            }

            return GetNameCase(caseNum);
        }

        /// <summary>
        ///     Выполняет склонение имени
        /// </summary>
        /// <param name="name">Имя</param>
        /// <param name="caseNum">Падеж</param>
        /// <returns>Имя в указанном падеже</returns>
        public string QName(string name, Padeg caseNum)
        {
            return QName(name, caseNum, Gender.Null);
        }

        /// <summary>
        ///     Выполняет склонение фамилии
        /// </summary>
        /// <param name="surname">Фамилия</param>
        /// <param name="gender">Пол</param>
        /// <returns>Массив со всеми падежами</returns>
        public string[] QSurname(string surname, Gender gender)
        {
            FullReset();
            SetSurname(surname);

            if (gender != Gender.Null)
            {
                SetGender(gender);
            }

            return GetSurnameCases();
        }

        /// <summary>
        ///     Выполняет склонение фамилии
        /// </summary>
        /// <param name="surname">Фамилия</param>
        /// <returns>Массив со всеми падежами</returns>
        public string[] QSurname(string surname)
        {
            return QSurname(surname, Gender.Null);
        }

        /// <summary>
        ///     Выполняет склонение фамилии
        /// </summary>
        /// <param name="surname">Фамилия</param>
        /// <param name="caseNum">Падеж</param>
        /// <param name="gender">Пол</param>
        /// <returns>Фамилия в указанном падеже</returns>
        public string QSurname(string surname, Padeg caseNum, Gender gender)
        {
            FullReset();
            SetSurname(surname);

            if (gender != Gender.Null)
            {
                SetGender(gender);
            }

            return GetSurnameCase(caseNum);
        }

        /// <summary>
        ///     Выполняет склонение фамилии
        /// </summary>
        /// <param name="surname">Фамилия</param>
        /// <param name="caseNum">Падеж</param>
        /// <returns>Фамилия в указанном падеже</returns>
        public string QSurname(string surname, Padeg caseNum)
        {
            return QSurname(surname, caseNum, Gender.Null);
        }

        /// <summary>
        ///     Выполняет склонение фамилии
        /// </summary>
        /// <param name="fatherName">Фамилия</param>
        /// <param name="gender">Пол</param>
        /// <returns>Массив со всеми падежами</returns>
        public string[] QFatherName(string fatherName, Gender gender)
        {
            FullReset();
            SetFatherName(fatherName);

            if (gender != Gender.Null)
            {
                SetGender(gender);
            }

            return GetFatherNameCases();
        }

        /// <summary>
        ///     Выполняет склонение фамилии
        /// </summary>
        /// <param name="fatherName">Фамилия</param>
        /// <returns>Массив со всеми падежами</returns>
        public string[] QFatherName(string fatherName)
        {
            return QFatherName(fatherName, Gender.Null);
        }

        /// <summary>
        ///     Выполняет склонение отчества
        /// </summary>
        /// <param name="fatherName">Отчество</param>
        /// <param name="caseNum">Падеж</param>
        /// <param name="gender">Пол</param>
        /// <returns>Отчество в указанном падеже</returns>
        public string QFatherName(string fatherName, Padeg caseNum, Gender gender)
        {
            FullReset();
            SetFatherName(fatherName);

            if (gender != Gender.Null)
            {
                SetGender(gender);
            }

            return GetFatherNameCase(caseNum);
        }

        /// <summary>
        ///     Выполняет склонение фамилии
        /// </summary>
        /// <param name="fatherName">Фамилия</param>
        /// <param name="caseNum">Падеж</param>
        /// <returns>Фамилия в указанном падеже</returns>
        public string QFatherName(string fatherName, Padeg caseNum)
        {
            return QFatherName(fatherName, caseNum, Gender.Null);
        }

        /// <summary>
        ///     Соединяет все слова которые есть в системе в одну строку в определенном падеже
        /// </summary>
        /// <param name="caseNum">Падеж</param>
        /// <returns>Соединенная строка</returns>
        private string ConnectedCase(Padeg caseNum)
        {
            var result = Words.Aggregate("", (current, w) => current + w.GetNameCase(caseNum) + " ");

            return result.TrimEnd();
        }

        /// <summary>
        ///     Соединяет все слова которые есть в системе в массив со всеми падежами
        /// </summary>
        /// <returns>Массив со всеми падежами</returns>
        private string[] ConnectedCases()
        {
            var res = new string[CaseCount];

            for (var i = 0; i < CaseCount; i++)
            {
                res[i] = ConnectedCase((Padeg)i);
            }

            return res;
        }

        /// <summary>
        ///     Выполняет склонение полного имени
        /// </summary>
        /// <param name="fullName">Полное имя</param>
        /// <param name="gender">Пол</param>
        /// <returns>Массив со всеми падежами</returns>
        public string[] QFullName(string fullName, Gender gender)
        {
            FullReset();
            SplitFullName(fullName);

            if (gender != Gender.Null)
            {
                SetGender(gender);
            }

            AllWordCases();

            return ConnectedCases();
        }

        /// <summary>
        ///     Выполняет склонение полного имени
        /// </summary>
        /// <param name="fullName">Полное имя</param>
        /// <returns>Массив со всеми падежами</returns>
        public string[] QFullName(string fullName)
        {
            return QFullName(fullName, Gender.Null);
        }

        /// <summary>
        ///     Выполняет склонение полного имени
        /// </summary>
        /// <param name="fullName">Полное имя</param>
        /// <param name="caseNum">Падеж</param>
        /// <param name="gender">Пол</param>
        /// <returns>Строка в указанном падеже</returns>
        public string QFullName(string fullName, Padeg caseNum, Gender gender)
        {
            FullReset();
            SplitFullName(fullName);

            if (gender != Gender.Null)
            {
                SetGender(gender);
            }

            AllWordCases();

            return ConnectedCase(caseNum);
        }

        /// <summary>
        ///     Выполняет склонение полного имени
        /// </summary>
        /// <param name="fullName">Полное имя</param>
        /// <param name="caseNum">Падеж</param>
        /// <returns>Строка в указанном падеже</returns>
        public string QFullName(string fullName, Padeg caseNum)
        {
            return QFullName(fullName, caseNum, Gender.Null);
        }

        /// <summary>
        ///     Определяет пол человека по ФИО
        /// </summary>
        /// <param name="fullName">ФИО</param>
        /// <returns>Пол человека</returns>
        public Gender DetectGender(string fullName)
        {
            FullReset();
            SplitFullName(fullName);

            return GenderAutoDetect();
        }

        /// <summary>
        ///     Массив содержит элементы типа Word. Это все слова которые нужно обработать и просклонять
        /// </summary>
        /// <returns>Массив всех слов</returns>
        public List<Word> Words { get; private set; }

        /// <summary>
        ///     Склонение имени по правилам мужских имен
        /// </summary>
        /// <returns>true если склонение было произведено и false если правило не было найденным</returns>
        protected abstract bool ManName();

        /// <summary>
        ///     Склонение имени по правилам женских имен
        /// </summary>
        /// <returns>true если склонение было произведено и false если правило не было найденным</returns>
        protected abstract bool WomanName();

        /// <summary>
        ///     Склонение фамилию по правилам мужских имен
        /// </summary>
        /// <returns>true если склонение было произведено и false если правило не было найденным</returns>
        protected abstract bool ManSurname();

        /// <summary>
        ///     Склонение фамилию по правилам женских имен
        /// </summary>
        /// <returns>true если склонение было произведено и false если правило не было найденным</returns>
        protected abstract bool WomanSurname();

        /// <summary>
        ///     Склонение отчества по правилам мужских имен
        /// </summary>
        /// <returns>true если склонение было произведено и false если правило не было найденным</returns>
        protected abstract bool ManFatherName();

        /// <summary>
        ///     Склонение отчества по правилам женских имен
        /// </summary>
        /// <returns>true если склонение было произведено и false если правило не было найденным</returns>
        protected abstract bool WomanFatherName();

        /// <summary>
        ///     Определяет пол человека по его имени
        /// </summary>
        /// <param name="word">Имя</param>
        protected abstract void GenderByName(Word word);

        /// <summary>
        ///     Определяет пол человека по его фамилии
        /// </summary>
        /// <param name="word">Фамилия</param>
        protected abstract void GenderBySurname(Word word);

        /// <summary>
        ///     Определяет пол человека по его отчеству
        /// </summary>
        /// <param name="word">Отчество</param>
        protected abstract void GenderByFatherName(Word word);

        /// <summary>
        ///     Идентифицирует слово определяя имя это, или фамилия, или отчество
        /// </summary>
        /// <param name="word">Слово для которое нужно идентифицировать</param>
        protected abstract void DetectNamePart(Word word);
    }
}