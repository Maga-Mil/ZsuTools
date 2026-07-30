#region

using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using NameCaseLib.Core;
using NameCaseLib.NCL;

#endregion

namespace NameCaseLib
{
    /// <summary>
    ///     Русские правила склонения ФИО
    ///     Правила определения пола человека по ФИО для русского языка
    ///     Система разделения фамилий имен и отчеств для русского языка
    /// </summary>
    public class Ru : Core.Core
    {
        /// <summary>
        ///     Версия языкового файла
        /// </summary>
        public override string LanguageVersion => "11072716";

        /// <summary>
        ///     Список согласных русского языка
        /// </summary>
        private const string Consonant = "бвгджзйклмнпрстфхцчшщ";

        /// <summary>
        ///     Окончания имен/фамилий, который не склоняются
        /// </summary>
        private readonly string[] ih = { "их", "ых", "ко", "уа" };

        /// <summary>
        ///     Окончания имен/фамилий, который не склоняются
        /// </summary>
        private readonly string[] ovo = { "ово", "аго", "яго", "ирь" };

        private readonly Dictionary<string, string> splitSecondExclude = new Dictionary<string, string>
        {
            { "а", "взйкмнпрстфя" },
            { "б", "а" },
            { "в", "аь" },
            { "г", "а" },
            { "д", "ар" },
            { "е", "бвгдйлмня" },
            { "ё", "бвгдйлмня" },
            { "ж", "" },
            { "з", "а" },
            { "и", "гдйклмнопрсфя" },
            { "й", "ля" },
            { "к", "аст" },
            { "л", "аилоья" },
            { "м", "аип" },
            { "н", "ат" },
            { "о", "вдлнпря" },
            { "п", "п" },
            { "р", "адикпть" },
            { "с", "атуя" },
            { "т", "аор" },
            { "у", "дмр" },
            { "ф", "аь" },
            { "х", "а" },
            { "ц", "а" },
            { "ч", "" },
            { "ш", "а" },
            { "щ", "" },
            { "ъ", "" },
            { "ы", "дн" },
            { "ь", "я" },
            { "э", "" },
            { "ю", "" },
            { "я", "нс" }
        };

        private readonly string[] namesMan =
        {
            "Вова", "Анри", "Питер", "Пауль", "Франц", "Вильям", "Уильям",
            "Альфонс", "Ганс", "Франс", "Филиппо", "Андреа", "Корнелис", "Фрэнк", "Леонардо",
            "Джеймс", "Отто", "Жан-пьер", "Джованни", "Джозеф", "Педро", "Адольф", "Уолтер",
            "Антонио", "Якоб", "Эсташ", "Адрианс", "Франческо", "Доменико", "Ханс", "Гун",
            "Шарль", "Хендрик", "Амброзиус", "Таддео", "Фердинанд", "Джошуа", "Изак", "Иоганн",
            "Фридрих", "Эмиль", "Умберто", "Франсуа", "Ян", "Эрнст", "Георг", "Карл"
        };

        /// <summary>
        ///     Список гласных русского языка
        /// </summary>
        private const string Vowels = "аеёиоуыэюя";

        /// <summary>
        ///     Количество падежей в языке
        /// </summary>
        protected override int CaseCount => 6;

        private readonly string[] foreignFatherNameParts = new string[]
            { "кизи", "кызы", "кызи", "кизы", "оглы", "огли", "угли", "углы", "уулу" };

        /// <summary>
        ///     Склоняет слово по нужным правилам
        /// </summary>
        /// <param name="word">Слово</param>
        protected override void WordCase(Word word)
        {
            if (word.NamePart == NamePart.FatherName && foreignFatherNameParts.Any(fnp => word.Name.Contains(fnp)))
            {
                var currentWords = word.WordOrig.Split(' ');

                if (currentWords.Length > 1)
                {
                    var result = new string[CaseCount];
                    var lastRule = -1;
                    var oCurWords = new List<Word>(currentWords.Length);

                    foreach (var w in currentWords)
                    {
                        var oNcw = new Word(w);

                        if (!foreignFatherNameParts.Contains(w, StringComparer.OrdinalIgnoreCase))
                        {
                            var instance = new Ru();
                            instance.DetectNamePart(oNcw);

                            SetWorkingWord(w);

                            string[] tmpResult;

                            if ((bool)GetType()
                                .GetMethod($"{word.Gender.ToString("g")}{oNcw.NamePart.ToString("g")}",
                                    BindingFlags.NonPublic | BindingFlags.Instance).Invoke(this, null))
                            {
                                tmpResult = LastResult;
                                lastRule = LastRule;
                            }
                            else
                            {
                                tmpResult = Enumerable.Repeat(w, CaseCount).ToArray();
                                lastRule = -1;
                            }

                            oNcw.SetNameCases(tmpResult);
                            oCurWords.Add(oNcw);
                        }
                        else
                        {
                            oNcw.SetNameCases(Enumerable.Repeat(w, CaseCount).ToArray());
                            oCurWords.Add(oNcw);

                            lastRule = -1;
                        }
                    }

                    foreach (var nameCases in oCurWords.Select(oNcw => oNcw.NameCases))
                    {
                        for (var k = 0; k < nameCases.Length; k++)
                        {
                            var nameCase = nameCases[k];

                            if (!string.IsNullOrEmpty(result[k]))
                            {
                                result[k] = $"{result[k]} {nameCase}";
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

                    return;
                }
            }

            base.WordCase(word);
        }

        /// <summary>
        ///     Мужские имена, оканчивающиеся на любой ь и -й,
        ///     склоняются так же, как обычные существительные мужского рода
        /// </summary>
        /// <returns>если правило было задействовано и false если нет</returns>
        protected bool ManRule1()
        {
            if (InLetters(Last(1), "ьй"))
            {
                if (InNames(WorkingWord, "Дель"))
                {
                    LastRule = 101;
                    MakeResultTheSame();

                    return true;
                }

                if (Last(2, 1) != "и")
                {
                    WordForms(WorkingWord, 1, "я", "ю", "я", "ем", "е");
                    LastRule = 102;

                    return true;
                }
                else
                {
                    WordForms(WorkingWord, 1, "я", "ю", "я", "ем", "и");
                    LastRule = 103;

                    return true;
                }
            }

            return false;
        }

        /// <summary>
        ///     Мужские имена, оканчивающиеся на любой твердый согласный,
        ///     склоняются так же, как обычные существительные мужского рода
        /// </summary>
        /// <returns>если правило было задействовано и false если нет</returns>
        protected bool ManRule2()
        {
            if (InLetters(Last(1), Consonant))
            {
                if (InNames(WorkingWord, "Павел"))
                {
                    LastResult = new[] { "Павел", "Павла", "Павлу", "Павла", "Павлом", "Павле" };
                    LastRule = 201;

                    return true;
                }
                else if (InNames(WorkingWord, "Лев"))
                {
                    LastResult = new[] { "Лев", "Льва", "Льву", "Льва", "Львом", "Льве" };
                    LastRule = 202;

                    return true;
                }
                else if (InNames(WorkingWord, "ван"))
                {
                    LastRule = 203;
                    MakeResultTheSame();

                    return true;
                }
                else
                {
                    WordForms(WorkingWord, "а", "у", "а", "ом", "е");
                    LastRule = 204;

                    return true;
                }
            }

            return false;
        }

        /// <summary>
        ///     Мужские и женские имена, оканчивающиеся на -а, склоняются, как и любые
        ///     существительные с таким же окончанием
        ///     Мужские и женские имена, оканчивающиеся иа -я, -ья, -ия, -ея, независимо от языка,
        ///     из которого они происходят, склоняются как существительные с соответствующими окончаниями
        /// </summary>
        /// <returns>если правило было задействовано и false если нет</returns>
        protected bool ManRule3()
        {
            if (Last(1) == "а")
            {
                if (InNames(WorkingWord, "фра", "Дега", "Андреа", "Сёра", "Сера"))
                {
                    LastRule = 301;
                    MakeResultTheSame();

                    return true;
                }

                if (!InLetters(Last(2, 1), "кшгх"))
                {
                    WordForms(WorkingWord, 1, "ы", "е", "у", "ой", "е");
                    LastRule = 302;

                    return true;
                }
                else
                {
                    WordForms(WorkingWord, 1, "и", "е", "у", "ой", "е");
                    LastRule = 303;

                    return true;
                }
            }

            if (Last(1) == "я")
            {
                WordForms(WorkingWord, 1, "и", "е", "ю", "ей", "е");
                LastRule = 303;

                return true;
            }

            return false;
        }

        /// <summary>
        ///     Мужские фамилии, оканчивающиеся на -ь -й, склоняются так же,
        ///     как обычные существительные мужского рода
        /// </summary>
        /// <returns>если правило было задействовано и false если нет</returns>
        protected bool ManRule4()
        {
            if (InLetters(Last(1), "ьй"))
            {
                if (Last(3) == "бей")
                {
                    WordForms(WorkingWord, 2, "ья", "ью", "ья", "ьем", "ье");
                    LastRule = 400;

                    return true;
                }
                else if (Last(3, 1) == "а" || InLetters(Last(2, 1), "ел"))
                {
                    WordForms(WorkingWord, 1, "я", "ю", "я", "ем", "е");
                    LastRule = 401;

                    return true;
                }

                //Толстой -» ТолстЫм 
                else if (Last(2, 1) == "ы" || Last(3, 1) == "т")
                {
                    WordForms(WorkingWord, 2, "ого", "ому", "ого", "ым", "ом");
                    LastRule = 402;

                    return true;
                }

                //Лесничий
                else if (Last(3) == "чий")
                {
                    WordForms(WorkingWord, 2, "ьего", "ьему", "ьего", "ьим", "ьем");
                    LastRule = 403;

                    return true;
                }
                else if (!InLetters(Last(2, 1), Vowels) || InLetters(Last(2, 1), "ио"))
                {
                    WordForms(WorkingWord, 2, "ого", "ому", "ого", "им", "ом");
                    LastRule = 404;

                    return true;
                }
                else
                {
                    MakeResultTheSame();
                    LastRule = 405;

                    return true;
                }
            }

            return false;
        }

        /// <summary>
        ///     Мужские фамилии, оканчивающиеся на -к
        /// </summary>
        /// <returns>если правило было задействовано и false если нет</returns>
        protected bool ManRule5()
        {
            if (Last(1) == "к")
            {
                //Если перед слово на ок, то нужно убрать о
                if (Last(4) == "енок" || Last(4) == "ёнок")
                {
                    WordForms(WorkingWord, 2, "ка", "ку", "ка", "ком", "ке");
                    LastRule = 501;

                    return true;
                }

                if (Last(2, 1) == "е" && !Last(3, 1).Contains("р") && Last(3) != "чек") // Гудачек
                {
                    WordForms(WorkingWord, 2, "ька", "ьку", "ька", "ьком", "ьке");
                    LastRule = 502;

                    return true;
                }
                else
                {
                    WordForms(WorkingWord, "а", "у", "а", "ом", "е");
                    LastRule = 503;

                    return true;
                }
            }

            return false;
        }

        /// <summary>
        ///     Мужские фамилии на согласный выбираем ем/ом/ым
        /// </summary>
        /// <returns>если правило было задействовано и false если нет</returns>
        protected bool ManRule6()
        {
            if (Last(1) == "ч")
            {
                WordForms(WorkingWord, "а", "у", "а", "ем", "е");
                LastRule = 601;

                return true;
            }

            //е перед ц выпадает

            if (Last(2) == "ец")
            {
                WordForms(WorkingWord, 2, "ца", "цу", "ца", "цом", "це");
                LastRule = 604;

                return true;
            }

            if (InLetters(Last(1), "цсршмхт"))
            {
                WordForms(WorkingWord, "а", "у", "а", "ом", "е");
                LastRule = 602;

                return true;
            }

            if (InLetters(Last(1), Consonant))
            {
                WordForms(WorkingWord, "а", "у", "а", "ым", "е");
                LastRule = 603;

                return true;
            }

            return false;
        }

        /// <summary>
        ///     Мужские фамилии на -а -я
        /// </summary>
        /// <returns>если правило было задействовано и false если нет</returns>
        protected bool ManRule7()
        {
            if (Last(1) == "а")
            {
                if (InNames(WorkingWord, "да"))
                {
                    LastRule = 701;
                    MakeResultTheSame();

                    return true;
                }

                //Если основа на ш, то нужно и, ей
                if (Last(2, 1) == "ш")
                {
                    WordForms(WorkingWord, 1, "и", "е", "у", "ей", "е");
                    LastRule = 702;

                    return true;
                }
                else if (InLetters(Last(2, 1), "хкг"))
                {
                    WordForms(WorkingWord, 1, "и", "е", "у", "ой", "е");
                    LastRule = 703;

                    return true;
                }
                else
                {
                    WordForms(WorkingWord, 1, "ы", "е", "у", "ой", "е");
                    LastRule = 704;

                    return true;
                }
            }

            if (Last(1) == "я")
            {
                WordForms(WorkingWord, 2, "ой", "ой", "ую", "ой", "ой");
                LastRule = 705;

                return true;
            }

            return false;
        }

        /// <summary>
        ///     Не склоняются мужские фамилии
        /// </summary>
        /// <returns>если правило было задействовано и false если нет</returns>
        protected bool ManRule8()
        {
            if (InEndings(Last(3), ovo) || InEndings(Last(2), ih))
            {
                if (InNames(WorkingWord, "Рерих"))
                {
                    return false;
                }

                LastRule = 8;
                MakeResultTheSame();

                return true;
            }

            return false;
        }

        /// <summary>
        ///     Мужские и женские имена, оканчивающиеся на -а, склоняются,
        ///     как и любые существительные с таким же окончанием
        /// </summary>
        /// <returns>если правило было задействовано и false если нет</returns>
        protected bool WomanRule1()
        {
            if (Last(1) == "а" && Last(2, 1) != "и")
            {
                if (!InLetters(Last(2, 1), "шхкг"))
                {
                    WordForms(WorkingWord, 1, "ы", "е", "у", "ой", "е");
                    LastRule = 101;

                    return true;
                }
                else
                {
                    //ей после шипящего
                    if (Last(2, 1) == "ш")
                    {
                        WordForms(WorkingWord, 1, "и", "е", "у", "ей", "е");
                        LastRule = 102;

                        return true;
                    }

                    WordForms(WorkingWord, 1, "и", "е", "у", "ой", "е");
                    LastRule = 103;

                    return true;
                }
            }

            return false;
        }

        /// <summary>
        ///     Мужские и женские имена, оканчивающиеся иа -я, -ья, -ия, -ея, независимо от языка,
        ///     из которого они происходят, склоняются как существительные с соответствующими окончаниями
        /// </summary>
        /// <returns>если правило было задействовано и false если нет</returns>
        protected bool WomanRule2()
        {
            if (Last(1) == "я")
            {
                if (Last(2, 1) != "и")
                {
                    WordForms(WorkingWord, 1, "и", "е", "ю", "ей", "е");
                    LastRule = 201;

                    return true;
                }
                else
                {
                    WordForms(WorkingWord, 1, "и", "и", "ю", "ей", "и");
                    LastRule = 202;

                    return true;
                }
            }

            return false;
        }

        /// <summary>
        ///     Русские женские имена, оканчивающиеся на мягкий согласный, склоняются,
        ///     как существительные женского рода типа дочь, тень
        /// </summary>
        /// <returns>если правило было задействовано и false если нет</returns>
        protected bool WomanRule3()
        {
            if (Last(1) == "ь")
            {
                WordForms(WorkingWord, 1, "и", "и", "ь", "ью", "и");
                LastRule = 3;

                return true;
            }

            return false;
        }

        /// <summary>
        ///     Женские фамилии, оканчивающиеся на -а -я, склоняются,
        ///     как и любые существительные с таким же окончанием
        /// </summary>
        /// <returns>если правило было задействовано и false если нет</returns>
        protected bool WomanRule4()
        {
            if (Last(1) == "а")
            {
                if (InLetters(Last(2, 1), "гк"))
                {
                    WordForms(WorkingWord, 1, "и", "е", "у", "ой", "е");
                    LastRule = 401;

                    return true;
                }
                else if (InLetters(Last(2, 1), "ш"))
                {
                    WordForms(WorkingWord, 1, "и", "е", "у", "ей", "е");
                    LastRule = 402;

                    return true;
                }
                else
                {
                    WordForms(WorkingWord, 1, "ой", "ой", "у", "ой", "ой");
                    LastRule = 403;

                    return true;
                }
            }

            else if (Last(1) == "я")
            {
                WordForms(WorkingWord, 2, "ой", "ой", "ую", "ой", "ой");
                LastRule = 404;

                return true;
            }

            return false;
        }

        /// <summary>
        ///     Функция пытается применить цепочку правил для мужских имен
        ///     @return boolean true - если было использовано правило из списка, false - если правило не было найденным
        /// </summary>
        protected override bool ManName()
        {
            if (InNames(WorkingWord, "Старший", "Младший"))
            {
                WordForms(WorkingWord, 2, "его", "ему", "его", "им", "ем");

                return true;
            }

            if (InNames(WorkingWord, "Мариа"))
            {
                //Альфонс Мария Муха
                WordForms(WorkingWord, 1, "и", "и", "ю", "ей", "ии");

                return true;
            }

            return RulesChain(Gender.Man, new[] { 1, 2, 3 });
        }

        /// <summary>
        ///     Функция пытается применить цепочку правил для женских имен
        ///     @return boolean true - если было использовано правило из списка, false - если правило не было найденным
        /// </summary>
        protected override bool WomanName()
        {
            return RulesChain(Gender.Woman, new[] { 1, 2, 3 });
        }

        /// <summary>
        ///     Функция пытается применить цепочку правил для мужских фамилий
        ///     @return boolean true - если было использовано правило из списка, false - если правило не было найденным
        /// </summary>
        protected override bool ManSurname()
        {
            return RulesChain(Gender.Man, new[] { 8, 4, 5, 6, 7 });
        }

        /// <summary>
        ///     Функция пытается применить цепочку правил для женских фамилий
        ///     @return boolean true - если было использовано правило из списка, false - если правило не было найденным
        /// </summary>
        protected override bool WomanSurname()
        {
            return RulesChain(Gender.Woman, new[] { 4 });
        }

        /// <summary>
        ///     Функция склоняет мужские отчества
        ///     @return boolean true - если слово было успешно изменено, false - если не получилось этого сделать
        /// </summary>
        protected override bool ManFatherName()
        {
            //Проверяем действительно ли отчество
            if (InNames(WorkingWord, "Ильич"))
            {
                WordForms(WorkingWord, "а", "у", "а", "ом", "е");

                return true;
            }

            if (Last(2) == "ич")
            {
                WordForms(WorkingWord, "а", "у", "а", "ем", "е");

                return true;
            }

            return false;
        }

        /// <summary>
        ///     Функция склоняет женские отчества
        ///     @return boolean true - если слово было успешно изменено, false - если не получилось этого сделать
        /// </summary>
        protected override bool WomanFatherName()
        {
            //Проверяем действительно ли отчество
            if (Last(2) == "на")
            {
                WordForms(WorkingWord, 1, "ы", "е", "у", "ой", "е");

                return true;
            }

            return false;
        }

        /// <summary>
        ///     Определение пола по правилам имен
        ///     @param NCLNameCaseWord word объект класса слов, для которого нужно определить пол
        /// </summary>
        protected override void GenderByName(Word word)
        {
            SetWorkingWord(word.Name);

            var prob = new GenderProbability();

            //Попробуем выжать максимум из имени
            //Если имя заканчивается на й, то скорее всего мужчина

            if (Last(1) == "й")
            {
                prob.Man += 0.9f;
            }

            if (InEndings(Last(2), "он", "ов", "ав", "ам", "ол", "ан", "рд", "мп", "по", "до", "др", "рт"))
            {
                prob.Man += 0.3f;
            }

            if (InLetters(Last(1), Consonant))
            {
                prob.Man += 0.01f;
            }

            if (Last(1) == "ь")
            {
                prob.Man += 0.02f;
            }

            if (InEndings(Last(2), "вь", "фь", "ль", "на"))
            {
                prob.Woman += 0.1f;
            }

            if (InEndings(Last(2), "ла"))
            {
                prob.Woman += 0.04f;
            }

            if (InEndings(Last(2), "то", "ма"))
            {
                prob.Man += 0.01f;
            }

            if (InEndings(Last(3), "лья", "вва", "ока", "ука", "ита", "эль", "реа"))
            {
                prob.Man += 0.2f;
            }

            if (InEndings(Last(3), "има"))
            {
                prob.Woman += 0.15f;
            }

            if (InEndings(Last(3), "лия", "ния", "сия", "дра", "лла", "кла", "опа", "вия"))
            {
                prob.Woman += 0.5f;
            }

            if (InEndings(Last(4), "льда", "фира", "нина", "лита", "алья"))
            {
                prob.Woman += 0.5f;
            }

            if (InNames(WorkingWord, namesMan))
            {
                prob.Man += 10;
            }

            if (InNames(WorkingWord, "Бриджет", "Элизабет", "Маргарет", "Джанет", "Жаклин", "Эвелин", "Чулпан", "Лариса"))
            {
                prob.Woman += 10;
            }

            //Исключение для Берил Кук, которая женщина
            if (InNames(WorkingWord, "Берил"))
            {
                prob.Woman += 0.05f;
            }

            word.GenderProbability = prob;
        }

        /// <summary>
        ///     Определение пола по правилам фамилий
        ///     @param NCLNameCaseWord word объект класса слов, для которого нужно определить пол
        /// </summary>
        protected override void GenderBySurname(Word word)
        {
            SetWorkingWord(word.Name);

            var prob = new GenderProbability();

            if (InEndings(Last(2), "ов", "ин", "ев", "ий", "ёв", "ый", "ын", "ой"))
            {
                prob.Man += 0.4f;
            }

            if (InEndings(Last(3), "ова", "ина", "ева", "ёва", "ына", "мин"))
            {
                prob.Woman += 0.4f;
            }

            if (InEndings(Last(2), "ая"))
            {
                prob.Woman += 0.4f;
            }

            word.GenderProbability = prob;
        }

        /// <summary>
        ///     Определение пола по правилам отчеств
        ///     @param NCLNameCaseWord word объект класса слов, для которого нужно определить пол
        /// </summary>
        protected override void GenderByFatherName(Word word)
        {
            SetWorkingWord(word.Name);

            if (Last(2) == "ич")
            {
                word.GenderProbability = new GenderProbability(10, 0); // мужчина
            }

            if (Last(2) == "на")
            {
                word.GenderProbability = new GenderProbability(0, 12); // женщина
            }

            if (InNames(WorkingWord, "оглы", "огли", "угли", "углы", "уулу"))
            {
                word.GenderProbability = new GenderProbability(10, 0); // мужчина
            }

            if (InNames(WorkingWord, "кизи", "кызы", "кызи", "кизы"))
            {
                word.GenderProbability = new GenderProbability(0, 12); // женщина
            }

            if (word.GenderProbability == null)
            {
                word.GenderProbability = new GenderProbability();
            }
        }

        /// <summary>
        ///     Идентифицирует слово определяет имя это, или фамилия, или отчество
        ///     - <b>N</b> - имя
        ///     - <b>S</b> - фамилия
        ///     - <b>F</b> - отчество
        ///     @param NCLNameCaseWord word объект класса слов, который необходимо идентифицировать
        /// </summary>
        protected override void DetectNamePart(Word word)
        {
            var name = word.Name;
            var length = name.Length;
            SetWorkingWord(name);

            //Считаем вероятность
            float first = 0;
            float surname = 0;
            float father = 0;

            //если смахивает на отчество
            if (InEndings(Last(3), "вна", "чна", "вич", "ьич"))
            {
                father += 3;
            }

            if (InEndings(Last(2), "ша"))
            {
                first += 0.5f;
            }

            if (InEndings(Last(3), "эль"))
            {
                first += 0.5f;
            }

            // буквы на которые никогда не заканчиваются имена
            if (InLetters(Last(1), "еёжхцочшщъыэю"))
            {
                //Просто исключения
                if (InNames(name, "Мауриц"))
                {
                    first += 10;
                }
                else
                {
                    surname += 0.3f;
                }
            }

            // Используем массив характерных окончаний
            if (InLetters(Last(2, 1), Vowels + Consonant))
            {
                if (!InLetters(Last(1), splitSecondExclude[Last(2, 1)]))
                {
                    surname += 0.4f;
                }
            }

            // Сокращенные ласкательные имена типа Аня Галя и.т.д.
            if (Last(1) == "я" && InLetters(Last(3, 1), Vowels))
            {
                first += 0.5f;
            }

            // Не бывает имен с такими предпоследними буквами
            if (InLetters(Last(2, 1), "жчщъэю"))
            {
                surname += 0.3f;
            }

            // Слова на мягкий знак. Существует очень мало имен на мягкий знак. Все остальное фамилии
            if (Last(1) == "ь")
            {
                if (Last(3, 2) == "ел")
                {
                    first += 0.7f;
                }

                // Просто исключения
                else if (InNames(name, "Лазарь", "Игорь", "Любовь"))
                {
                    first += 10;
                }

                // Если не то и не другое, тогда фамилия
                else
                {
                    surname += 0.3f;
                }
            }

            // Если две последних буквы согласные то скорее всего это фамилия
            else if (InLetters(Last(1), Consonant + "ь") && InLetters(Last(2, 1), Consonant + "ь"))
            {
                if (!InEndings(Last(2), "др", "кт", "лл", "пп", "рд", "рк", "рп", "рт", "тр"))
                {
                    surname += 0.25f;
                }
            }

            // Слова, которые заканчиваются на тин
            if (Last(3) == "тин" && InLetters(Last(4, 1), "нст"))
            {
                first += 0.5f;
            }

            //Исключения
            if (InNames(name, "Лев", "Яков", "Маша", "Ольга", "Еремей", "Исак", "Исаак", "Ева", "Ирина", "Элькин",
                    "Мерлин", "Макс", "Алекс", "Мариа", "Бриджет", "Элизабет", "Маргарет", "Джанет", "Жаклин",
                    "Эвелин") ||
                InNames(name, namesMan))
            {
                first += 10;
            }

            // Фамилии которые заканчиваются на -ли кроме тех что типа натАли и.т.д.
            if (Last(2) == "ли" && Last(3, 1) != "а")
            {
                surname += 0.4f;
            }

            // Фамилии на -як кроме тех что типа Касьян Куприян + Ян и.т.д.
            if (Last(2) == "ян" && length > 2 && !InLetters(Last(3, 1), "ьи"))
            {
                surname += 0.4f;
            }

            // Фамилии на -ур кроме имен Артур Тимур
            if (Last(2) == "ур")
            {
                if (!InNames(name, "Артур", "Тимур"))
                {
                    surname += 0.4f;
                }
            }

            // Разбор ласкательных имен на -ик
            if (Last(2) == "ик")
            {
                if (InLetters(Last(3, 1), "лшхд"))
                {
                    first += 0.3f;
                }
                else
                {
                    surname += 0.4f;
                }
            }

            // Разбор имен и фамилий, который заканчиваются на ина
            if (Last(3) == "ина")
            {
                if (InEndings(Last(7), "атерина", "ристина"))
                {
                    first += 10;
                }

                // Исключения
                else if (InNames(name, "Мальвина", "Антонина", "Альбина", "Агриппина", "Фаина", "Карина", "Марина",
                    "Валентина", "Калина", "Аделина", "Алина", "Ангелина", "Галина", "Каролина", "Павлина", "Полина",
                    "Элина", "Мина", "Нина", "Дина"))
                {
                    first += 10;
                }

                // Иначе фамилия
                else
                {
                    surname += 0.4f;
                }
            }

            // Имена типа Николай
            if (Last(4) == "олай")
            {
                first += 0.6f;
            }

            // Фамильные окончания
            if (InEndings(Last(2), "ов", "ин", "ев", "ёв", "ый", "ын", "ой", "ук", "як", "ца", "ун", "ок", "ая", "ёк",
                "ив", "ус", "ак", "яр", "уз", "ах", "ай"))
            {
                surname += 0.4f;
            }

            if (InEndings(Last(3), "ова", "ева", "ёва", "ына", "шен", "мей", "вка", "шир", "бан", "чий", "кий", "бей",
                "чан", "ган", "ким", "кан", "мар", "лис"))
            {
                surname += 0.4f;
            }

            if (InEndings(Last(4), "шена"))
            {
                surname += 0.4f;
            }

            //исключения и частички
            if (InNames(name, "да", "валадон", "данбар"))
            {
                surname += 10;
            }

            if (InNames(name, foreignFatherNameParts))
            {
                father += 10;
            }

            var max = Math.Max(Math.Max(first, surname), father);

            if (first == max)
            {
                word.NamePart = NamePart.Name;
            }
            else if (surname == max)
            {
                word.NamePart = NamePart.Surname;
            }
            else
            {
                word.NamePart = NamePart.FatherName;
            }
        }
    }
}