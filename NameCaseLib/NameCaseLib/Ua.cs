#region

using System;
using NameCaseLib.Core;
using NameCaseLib.NCL;

#endregion

namespace NameCaseLib
{
    /// <summary>
    ///     Украинские правила склонений ФИО.
    ///     Правила определения пола человека по ФИО для украинского языка
    ///     Система разделения фамилий имен и отчеств для украинского языка
    /// </summary>
    public class Ua : Core.Core
    {
        /// <summary>
        ///     Список согласных украинского языка
        /// </summary>
        private readonly string consonant = "бвгджзйклмнпрстфхцчшщ";

        /// <summary>
        ///     Українські губні звуки
        /// </summary>
        private readonly string gubni = "мвпбф";

        /// <summary>
        ///     Українські завжди м’які звуки
        /// </summary>
        private readonly string myaki = "ьюяєї";

        /// <summary>
        ///     Українські нешиплячі приголосні
        /// </summary>
        private readonly string neshyplyachi = "бвгдзклмнпрстфхц";

        /// <summary>
        ///     Українські шиплячі приголосні
        /// </summary>
        private readonly string shyplyachi = "жчшщ";

        /// <summary>
        ///     Список гласных украинского языка
        /// </summary>
        private readonly string vowels = "аеиоуіїєюя";

        /// <summary>
        ///     Количество падежей в языке
        /// </summary>
        protected override int CaseCount => 7;

        /// <summary>
        ///     Версия языкового файла
        /// </summary>
        public override string LanguageVersion => "11071222";

        /// <summary>
        ///     Чергування українських приголосних
        ///     Чергування г к х —» з ц с
        ///     <param name="letter">літера, яку необхідно перевірити на чергування</param>
        /// </summary>
        /// <returns>літера, де вже відбулося чергування</returns>
        private string inverseGKH(string letter)
        {
            switch (letter)
            {
                case "г": return "з";
                case "к": return "ц";
                case "х": return "с";
            }

            return letter;
        }

        /// <summary>
        ///     Перевіряє чи символ є апострофом чи не є
        ///     <param name="letter">симпол для перевірки</param>
        /// </summary>
        /// <returns>true якщо символ є апострофом</returns>
        private bool isApostrof(string letter)
        {
            if (InLetters(letter, " " + consonant + vowels))
            {
                return false;
            }

            return true;
        }

        /// <summary>
        ///     Чергування українських приголосних
        ///     Чергування г к —» ж ч
        ///     @param string letter літера, яку необхідно перевірити на чергування
        ///     <returns>літера, де вже відбулося чергування</returns>
        /// </summary>
        private string inverse2(string letter)
        {
            switch (letter)
            {
                case "к": return "ч";
                case "г": return "ж";
            }

            return letter;
        }

        /// <summary>
        ///     <b>Визначення групи для іменників 2-ї відміни</b>
        ///     1 - тверда
        ///     2 - мішана
        ///     3 - м’яка
        ///     <b>Правило:</b>
        ///     - Іменники з основою на твердий нешиплячий належать до твердої групи:
        ///     береза, дорога, Дніпро, шлях, віз, село, яблуко.
        ///     - Іменники з основою на твердий шиплячий належать до мішаної групи:
        ///     пожеж-а, пущ-а, тиш-а, алич-а, вуж, кущ, плющ, ключ, плече, прізвище.
        ///     - Іменники з основою на будь-який м"який чи пом"якше­ний належать до м"якої групи:
        ///     земля [земл"а], зоря [зор"а], армія [арм"ійа], сім"я [с"імйа], серпень, фахівець,
        ///     трамвай, су­зір"я [суз"ірйа], насіння [насін"н"а], узвишшя Іузвиш"ш"а
        ///     <param name="word">іменник, групу якого необхідно визначити</param>
        /// </summary>
        /// <returns>номер групи іменника</returns>
        private int detect2Group(string word)
        {
            var osnova = word;
            var stack = "";

            //Ріжемо слово поки не зустрінемо приголосний і записуемо в стек всі голосні які зустріли
            while (InLetters(Last(osnova, 1), vowels + "ь"))
            {
                stack = Last(osnova, 1) + stack;
                osnova = osnova.Substring(0, osnova.Length - 1);
            }

            var stacksize = stack.Length;
            var last = "Z"; //нульове закінчення

            if (stacksize > 0)
            {
                last = stack.Substring(0, 1);
            }

            var osnovaEnd = Last(osnova, 1);

            if (InLetters(osnovaEnd, neshyplyachi) && !InLetters(last, myaki))
            {
                return 1;
            }

            if (InLetters(osnovaEnd, shyplyachi) && !InLetters(last, myaki))
            {
                return 2;
            }

            return 3;
        }

        /// <summary>
        ///     Шукаємо першу з кінця літеру з переліку
        /// </summary>
        /// <param name="word">Слово</param>
        /// <param name="vowels">Перелік літер</param>
        /// <returns>Перша літера з кінця</returns>
        private string FirstLastVowel(string word, string vowels)
        {
            var length = word.Length;

            for (var i = length - 1; i > 0; i--)
            {
                var letter = word.Substring(i, 1);

                if (InLetters(letter, vowels))
                {
                    return letter;
                }
            }

            return "";
        }

        /// <summary>
        ///     Отримуємо основу слова за правилами української мови
        /// </summary>
        /// <param name="word">Слово</param>
        /// <returns>Основа слова</returns>
        private string getOsnova(string word)
        {
            var osnova = word;

            //Ріжемо слово поки не зустрінемо приголосний
            while (InLetters(Last(osnova, 1), vowels + "ь"))
            {
                osnova = osnova.Substring(0, osnova.Length - 1);
            }

            return osnova;
        }

        /// <summary>
        ///     Українські чоловічі та жіночі імена, що в називному відмінку однини закінчуються на -а (-я),
        ///     відмінються як відповідні іменники І відміни.
        ///     <ul>
        ///         <li>
        ///             Примітка 1. Кінцеві приголосні основи г, к, х у жіночих іменах
        ///             у давальному та місцевому відмінках однини перед закінченням -і
        ///             змінюються на з, ц, с: Ольга - Ользі, Палажка - Палажці, Солоха - Солосі.
        ///         </li>
        ///         <li>
        ///             Примітка 2. У жіночих іменах типу Одарка, Параска в родовому відмінку множини
        ///             в кінці основи між приголосними з"являється звук о: Одарок, Парасок.
        ///         </li>
        ///     </ul>
        /// </summary>
        /// <returns>true - якщо було задіяно правило з переліку, false - якщо правило не знайдено</returns>
        protected bool ManRule1()
        {
            //Предпоследний символ
            var beforeLast = Last(2, 1);

            //Останні літера або а
            if (Last(1) == "а")
            {
                WordForms(WorkingWord, 2, beforeLast + "и", inverseGKH(beforeLast) + "і", beforeLast + "у",
                    beforeLast + "ою", inverseGKH(beforeLast) + "і", beforeLast + "о");

                LastRule = 101;

                return true;
            }

            //Остання літера я

            if (Last(1) == "я")
            {
                if (beforeLast == "і")
                {
                    WordForms(WorkingWord, 1, "ї", "ї", "ю", "єю", "ї", "є");
                    LastRule = 102;

                    return true;
                }
                else
                {
                    WordForms(WorkingWord, 2, beforeLast + "і", inverseGKH(beforeLast) + "і", beforeLast + "ю",
                        beforeLast + "ею", inverseGKH(beforeLast) + "і", beforeLast + "е");

                    LastRule = 103;

                    return true;
                }
            }

            return false;
        }

        /// <summary>
        ///     Імена, що в називному відмінку закінчуються на -р, у родовому мають закінчення -а:
        ///     Віктор - Віктора, Макар - Макара, але: Ігор - Ігоря, Лазар - Лазаря.
        /// </summary>
        /// <returns>true - якщо було задіяно правило з переліку, false - якщо правило не знайдено</returns>
        protected bool ManRule2()
        {
            if (Last(1) == "р")
            {
                if (InNames(WorkingWord, "ігор", "лазар"))
                {
                    WordForms(WorkingWord, "я", "еві", "я", "ем", "еві", "е");
                    LastRule = 201;

                    return true;
                }
                else
                {
                    var osnova = WorkingWord;

                    if (Last(osnova, 2, 1) == "і")
                    {
                        osnova = osnova.Substring(0, osnova.Length - 2) + "о" + Last(osnova, 1);
                    }

                    WordForms(osnova, "а", "ові", "а", "ом", "ові", "е");
                    LastRule = 202;

                    return true;
                }
            }

            return false;
        }

        /// <summary>
        ///     Українські чоловічі імена, що в називному відмінку однини закінчуються на приголосний та -о,
        ///     відмінюються як відповідні іменники ІІ відміни.
        /// </summary>
        /// <returns>true - якщо було задіяно правило з переліку, false - якщо правило не знайдено</returns>
        protected bool ManRule3()
        {
            //Предпоследний символ
            var beforeLast = Last(2, 1);

            if (InLetters(Last(1), consonant + "оь"))
            {
                var group = detect2Group(WorkingWord);
                var osnova = getOsnova(WorkingWord);

                //В іменах типу Антін, Нестір, Нечипір, Прокіп, Сидір, Тиміш, Федір голосний і виступає тільки в 
                //називному відмінку, у непрямих - о: Антона, Антонові                           
                //Чергування і -» о всередині
                var osLast = Last(osnova, 1);

                if (osLast != "й" && Last(osnova, 2, 1) == "і" && !InEndings(Last(osnova, 4), "світ", "цвіт") &&
                    !InNames(WorkingWord, "гліб") &&
                    !InEndings(Last(2), "ік", "іч"))
                {
                    osnova = osnova.Substring(0, osnova.Length - 2) + "о" + Last(osnova, 1);
                }

                //Випадання букви е при відмінюванні слів типу Орел
                if (osnova != "" && osnova.Substring(0, 1) == "о" && FirstLastVowel(osnova, vowels + "гк") == "е" &&
                    Last(2) != "сь")
                {
                    var delim = osnova.LastIndexOf("е");
                    osnova = osnova.Substring(0, delim) + osnova.Substring(delim + 1, osnova.Length - delim);
                }

                if (group == 1)
                {
                    if (Last(2) == "ок" && Last(3) != "оок")
                    {
                        WordForms(WorkingWord, 2, "ка", "кові", "ка", "ком", "кові", "че");
                        LastRule = 301;

                        return true;
                    }

                    //Російські прізвища на ов, ев, єв
                    else if (InEndings(Last(2), "ов", "ев", "єв") &&
                             !InNames(WorkingWord, "лев", "остромов"))
                    {
                        WordForms(osnova, 1, osLast + "а", osLast + "у", osLast + "а", osLast + "им", osLast + "у",
                            inverse2(osLast) + "е");

                        LastRule = 302;

                        return true;
                    }

                    //Російські прізвища на ін
                    else if (InEndings(Last(2), "ін"))
                    {
                        WordForms(WorkingWord, "а", "у", "а", "ом", "у", "е");
                        LastRule = 303;

                        return true;
                    }
                    else
                    {
                        WordForms(osnova, 1, osLast + "а", osLast + "ові", osLast + "а", osLast + "ом", osLast + "ові",
                            inverse2(osLast) + "е");

                        LastRule = 304;

                        return true;
                    }
                }

                if (group == 2)
                {
                    //Мішана група
                    WordForms(osnova, "а", "еві", "а", "ем", "еві", "е");
                    LastRule = 305;

                    return true;
                }

                if (group == 3)
                {
                    if (Last(2) == "ей" && InLetters(Last(3, 1), gubni))
                    {
                        osnova = WorkingWord.Substring(0, WorkingWord.Length - 2) + "’";
                        WordForms(osnova, "я", "єві", "я", "єм", "єві", "ю");
                        LastRule = 306;

                        return true;
                    }
                    else if (Last(1) == "й" || beforeLast == "і")
                    {
                        WordForms(WorkingWord, 1, "я", "єві", "я", "єм", "єві", "ю");
                        LastRule = 307;

                        return true;
                    }

                    //Швець
                    else if (WorkingWord == "швець")
                    {
                        WordForms(WorkingWord, 4, "евця", "евцеві", "евця", "евцем", "евцеві", "евцю");
                        LastRule = 308;

                        return true;
                    }

                    //Слова що закінчуються на ець
                    else if (Last(3) == "ець")
                    {
                        WordForms(WorkingWord, 3, "ця", "цеві", "ця", "цем", "цеві", "цю");
                        LastRule = 309;

                        return true;
                    }

                    //Слова що закінчуються на єць яць
                    else if (InEndings(Last(3), "єць", "яць"))
                    {
                        WordForms(WorkingWord, 3, "йця", "йцеві", "йця", "йцем", "йцеві", "йцю");
                        LastRule = 310;

                        return true;
                    }
                    else
                    {
                        WordForms(osnova, "я", "еві", "я", "ем", "еві", "ю");
                        LastRule = 311;

                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        ///     Якщо слово закінчується на і, то відмінюємо як множину
        /// </summary>
        /// <returns>true - якщо було задіяно правило з переліку, false - якщо правило не знайдено</returns>
        protected bool ManRule4()
        {
            if (Last(1) == "і")
            {
                WordForms(WorkingWord, 1, "их", "им", "их", "ими", "их", "і");
                LastRule = 4;

                return true;
            }

            return false;
        }

        /// <summary>
        ///     Якщо слово закінчується на ий або ой
        /// </summary>
        /// <returns>true - якщо було задіяно правило з переліку, false - якщо правило не знайдено</returns>
        protected bool ManRule5()
        {
            if (InEndings(Last(2), "ий", "ой"))
            {
                WordForms(WorkingWord, 2, "ого", "ому", "ого", "им", "ому", "ий");
                LastRule = 5;

                return true;
            }

            return false;
        }

        /// <summary>
        ///     Українські чоловічі та жіночі імена, що в називному відмінку однини закінчуються на -а (-я),
        ///     відмінються як відповідні іменники І відміни.
        ///     - Примітка 1. Кінцеві приголосні основи г, к, х у жіночих іменах
        ///     у давальному та місцевому відмінках однини перед закінченням -і
        ///     змінюються на з, ц, с: Ольга - Ользі, Палажка - Палажці, Солоха - Солосі.
        ///     - Примітка 2. У жіночих іменах типу Одарка, Параска в родовому відмінку множини
        ///     в кінці основи між приголосними з"являється звук о: Одарок, Парасок
        /// </summary>
        /// <returns>true - якщо було задіяно правило з переліку, false - якщо правило не знайдено</returns>
        protected bool WomanRule1()
        {
            //Предпоследний символ
            var beforeLast = Last(2, 1);

            //Якщо закінчується на ніга -» нога
            if (Last(4) == "ніга")
            {
                var osnova = WorkingWord.Substring(0, WorkingWord.Length - 3) + "о";
                WordForms(osnova, "ги", "зі", "гу", "гою", "зі", "го");
                LastRule = 101;

                return true;
            }

            //Останні літера або а

            if (Last(1) == "а")
            {
                WordForms(WorkingWord, 2, beforeLast + "и", inverseGKH(beforeLast) + "і", beforeLast + "у",
                    beforeLast + "ою", inverseGKH(beforeLast) + "і", beforeLast + "о");

                LastRule = 102;

                return true;
            }

            //Остання літера я

            if (Last(1) == "я")
            {
                if (InLetters(beforeLast, vowels) || isApostrof(beforeLast))
                {
                    WordForms(WorkingWord, 1, "ї", "ї", "ю", "єю", "ї", "є");
                    LastRule = 103;

                    return true;
                }
                else
                {
                    WordForms(WorkingWord, 2, beforeLast + "і", inverseGKH(beforeLast) + "і", beforeLast + "ю",
                        beforeLast + "ею", inverseGKH(beforeLast) + "і", beforeLast + "е");

                    LastRule = 104;

                    return true;
                }
            }

            return false;
        }

        /// <summary>
        ///     Українські жіночі імена, що в називному відмінку однини закінчуються на приголосний,
        ///     відмінюються як відповідні іменники ІІІ відміни
        /// </summary>
        /// <returns>true - якщо було задіяно правило з переліку, false - якщо правило не знайдено</returns>
        protected bool WomanRule2()
        {
            if (InLetters(Last(1), consonant + "ь"))
            {
                var osnova = getOsnova(WorkingWord);
                var apostrof = "";
                var duplicate = "";
                var osLast = Last(osnova, 1);
                var osbeforeLast = Last(osnova, 2, 1);

                //Чи треба ставити апостроф
                if (InLetters(osLast, "мвпбф") && InLetters(osbeforeLast, vowels))
                {
                    apostrof = "’";
                }

                //Чи треба подвоювати
                if (InLetters(osLast, "дтзсцлн"))
                {
                    duplicate = osLast;
                }

                //Відмінюємо
                if (Last(1) == "ь")
                {
                    WordForms(osnova, "і", "і", "ь", duplicate + apostrof + "ю", "і", "е");
                    LastRule = 201;

                    return true;
                }

                WordForms(osnova, "і", "і", "", duplicate + apostrof + "ю", "і", "е");
                LastRule = 202;

                return true;
            }

            return false;
        }

        /// <summary>
        ///     Якщо слово на ськ або це російське прізвище
        /// </summary>
        /// <returns>true - якщо було задіяно правило з переліку, false - якщо правило не знайдено</returns>
        protected bool WomanRule3()
        {
            //Предпоследний символ
            var beforeLast = Last(2, 1);

            //Донская
            if (Last(2) == "ая")
            {
                WordForms(WorkingWord, 2, "ої", "ій", "ую", "ою", "ій", "ая");
                LastRule = 301;

                return true;
            }

            //Ті що на ськ
            if (Last(1) == "а" && (InLetters(Last(2, 1), "чнв") || InEndings(Last(3, 2), "ьк")))
            {
                WordForms(WorkingWord, 2, beforeLast + "ої", beforeLast + "ій", beforeLast + "у", beforeLast + "ою",
                    beforeLast + "ій", beforeLast + "о");

                LastRule = 302;

                return true;
            }

            return false;
        }

        /// <summary>
        ///     Функція намагається застосувати ланцюг правил для чоловічих імен
        /// </summary>
        /// <returns>true - якщо було задіяно правило з переліку, false - якщо правило не знайдено</returns>
        protected override bool ManName()
        {
            return RulesChain(Gender.Man, new[] { 1, 2, 3 });
        }

        /// <summary>
        ///     Функція намагається застосувати ланцюг правил для жіночих імен
        /// </summary>
        /// <returns>true - якщо було задіяно правило з переліку, false - якщо правило не знайдено</returns>
        protected override bool WomanName()
        {
            return RulesChain(Gender.Woman, new[] { 1, 2 });
        }

        /// <summary>
        ///     Функція намагається застосувати ланцюг правил для чоловічих прізвищ
        /// </summary>
        /// <returns>true - якщо було задіяно правило з переліку, false - якщо правило не знайдено</returns>
        protected override bool ManSurname()
        {
            return RulesChain(Gender.Man, new[] { 5, 1, 2, 3, 4 });
        }

        /// <summary>
        ///     Функція намагається застосувати ланцюг правил для жіночих прізвищ
        /// </summary>
        /// <returns>true - якщо було задіяно правило з переліку, false - якщо правило не знайдено</returns>
        protected override bool WomanSurname()
        {
            return RulesChain(Gender.Woman, new[] { 3, 1 });
        }

        /// <summary>
        ///     Фунція відмінює чоловічі по-батькові
        /// </summary>
        /// <returns>true - якщо слово успішно змінене, false - якщо невдалося провідміняти слово</returns>
        protected override bool ManFatherName()
        {
            if (InEndings(Last(2), "ич", "іч"))
            {
                WordForms(WorkingWord, "а", "у", "а", "ем", "у", "у");

                return true;
            }

            return false;
        }

        /// <summary>
        ///     Фунція відмінює жіночі по-батькові
        /// </summary>
        /// <returns>true - якщо слово успішно змінене, false - якщо невдалося провідміняти слово</returns>
        protected override bool WomanFatherName()
        {
            if (InEndings(Last(3), "вна"))
            {
                WordForms(WorkingWord, 1, "и", "і", "у", "ою", "і", "о");

                return true;
            }

            return false;
        }

        /// <summary>
        ///     Визначення статі, за правилами імені
        ///     <param name="word">Слово</param>
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

            if (InNames(WorkingWord, "петро", "микола"))
            {
                prob.Man += 30;
            }

            if (InEndings(Last(2), "он", "ов", "ав", "ам", "ол", "ан", "рд", "мп", "ко", "ло"))
            {
                prob.Man += 0.5f;
            }

            if (InEndings(Last(3), "бов", "нка", "яра", "ила", "опа"))
            {
                prob.Woman += 0.5f;
            }

            if (InLetters(Last(1), consonant))
            {
                prob.Man += 0.01f;
            }

            if (Last(1) == "ь")
            {
                prob.Man += 0.02f;
            }

            if (InEndings(Last(2), "дь"))
            {
                prob.Woman += 0.1f;
            }

            if (InEndings(Last(3), "ель", "бов"))
            {
                prob.Woman += 0.4f;
            }

            word.GenderProbability = prob;
        }

        /// <summary>
        ///     Визначення статі, за правилами прізвища
        ///     <param name="word">Слово</param>
        /// </summary>
        protected override void GenderBySurname(Word word)
        {
            SetWorkingWord(word.Name);

            var prob = new GenderProbability();

            if (InEndings(Last(2), "ов", "ин", "ев", "єв", "ін", "їн", "ий", "їв", "ів", "ой", "ей"))
            {
                prob.Man += 0.4f;
            }

            if (InEndings(Last(3), "ова", "ина", "ева", "єва", "іна", "мін"))
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
        ///     Визначення статі, за правилами по-батькові
        ///     <param name="word">Слово</param>
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
            
            if (word.GenderProbability == null)
            {
                word.GenderProbability = new GenderProbability();
            }
        }

        /// <summary>
        ///     Ідентифікує слово визначаючи чи це ім’я, чи це прізвище, чи це побатькові
        ///     <param name="word">Слово</param>
        /// </summary>
        protected override void DetectNamePart(Word word)
        {
            var namepart = word.Name;
            SetWorkingWord(namepart);

            //Считаем вероятность
            float first = 0;
            float second = 0;
            float father = 0;

            //если смахивает на отчество
            if (InEndings(Last(3), "вна", "чна", "ліч") || InEndings(Last(4), "ьмич", "ович"))
            {
                father += 3;
            }

            //Похоже на имя
            if (InEndings(Last(3), "тин") || InEndings(Last(4), "ьмич", "юбов", "івна", "явка", "орив", "кіян"))
            {
                first += 0.5f;
            }

            //Исключения
            if (InNames(namepart, "лев", "гаїна", "афіна", "антоніна", "ангеліна", "альвіна", "альбіна", "аліна",
                "павло", "олесь", "микола", "мая", "англеліна", "елькін", "мерлін"))
            {
                first += 10;
            }

            //похоже на фамилию
            if (InEndings(Last(2), "ов", "ін", "ев", "єв", "ий", "ин", "ой", "ко", "ук", "як", "ца", "их", "ик", "ун",
                "ок", "ша", "ая", "га", "єк", "аш", "ив", "юк", "ус", "це", "ак", "бр", "яр", "іл", "ів", "ич", "сь",
                "ей", "нс", "яс", "ер", "ай", "ян", "ах", "ць", "ющ", "іс", "ач", "уб", "ох", "юх", "ут", "ча", "ул",
                "вк", "зь", "уц", "їн", "де", "уз", "юр", "ік", "іч", "ро"))
            {
                second += 0.4f;
            }

            if (InEndings(Last(3), "ова", "ева", "єва", "тих", "рик", "вач", "аха", "шен", "мей", "арь", "вка", "шир",
                "бан", "чий", "іна", "їна", "ька", "ань", "ива", "аль", "ура", "ран", "ало", "ола", "кур", "оба", "оль",
                "нта", "зій", "ґан", "іло", "шта", "юпа", "рна", "бла", "еїн", "има", "мар", "кар", "оха", "чур", "ниш",
                "ета", "тна", "зур", "нір", "йма", "орж", "рба", "іла", "лас", "дід", "роз", "аба", "чан", "ган"))
            {
                second += 0.4f;
            }

            if (InEndings(Last(4), "ьник", "нчук", "тник", "кирь", "ский", "шена", "шина", "вина", "нина", "гана",
                "гана", "хній", "зюба", "орош", "орон", "сило", "руба", "лест", "мара", "обка", "рока", "сика", "одна",
                "нчар", "вата", "ндар", "грій"))
            {
                second += 0.4f;
            }

            if (Last(1) == "і")
            {
                second += 0.2f;
            }

            var max = Math.Max(Math.Max(first, second), father);

            if (first == max)
            {
                word.NamePart = NamePart.Name;
            }
            else if (second == max)
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