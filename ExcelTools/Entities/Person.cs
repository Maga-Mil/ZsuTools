using System;

namespace ZsuTools.Entities
{
    public class Person : IEquatable<Person>
    {
        /// <summary>
        /// Ім'я особи (First Name)
        /// </summary>
        public readonly string FirstName;
        /// <summary>
        /// По-батькові особи (Patronymic)
        /// </summary>
        public readonly string Patronymic;
        /// <summary>
        /// Прізвище особи (Last Name)
        /// </summary>
        public readonly string LastName;

        public Person(string stringToParse)
        {
            // Розділяємо рядок на слова
            var parts = stringToParse.Split(new[] { ' ', '.' }, StringSplitOptions.RemoveEmptyEntries);

            // 3 слівний формат: КОВАЛЕНКО Іван Петрович
            if (parts.Length == 3)
            {
                FirstName = Utils.FirstLetterUppercase(parts[1].Trim());
                Patronymic = Utils.FirstLetterUppercase(parts[2].Trim());
                LastName = parts[0].ToUpperInvariant(); //Згідно інструкції з діловодства
            }
            else if(parts.Length == 2)
            {
                // Can be Name LASTNAME format. Lastname must be in uppercase
                if(parts[1] == parts[1].ToUpperInvariant() && parts[1].Length > 1)
                {
                    LastName = parts[1].Trim();
                    FirstName = Utils.FirstLetterUppercase(parts[0].Trim());
                    Patronymic = string.Empty;
                }
                else if (parts[0] == parts[0].ToUpperInvariant() && parts[0].Length > 1)
                {
                    LastName = parts[0].Trim();
                    FirstName = Utils.FirstLetterUppercase(parts[1].Trim());
                    Patronymic = string.Empty;
                }
                //TODO consider support "Іван Іванович" format
                else
                {
                    throw new ArgumentException($"Cannot parse the name string {stringToParse}");
                }
            }
            else
            {
                throw new ArgumentException($"Cannot parse the name string {stringToParse}");
            }
        }
        
        public Person(string firstName, string patronymic, string lastName)
        {
            FirstName = Utils.FirstLetterUppercase(firstName);
            Patronymic = Utils.FirstLetterUppercase(patronymic);
            LastName = lastName.ToUpperInvariant(); //Згідно інструкції з діловодства
        }

        public Person(string firstName, string lastName)
        {
            FirstName = Utils.FirstLetterUppercase(firstName);
            Patronymic = string.Empty;
            LastName = lastName.ToUpperInvariant(); //Згідно інструкції з діловодства
        }

        public bool Equals(Person other)
        {
            if (other is null) return false;
            if (ReferenceEquals(this, other)) return true;
            var mandatoryCheck = string.Equals(FirstName, other.FirstName, StringComparison.InvariantCulture) && string.Equals(LastName, other.LastName, StringComparison.InvariantCulture);
            if (mandatoryCheck)
            {
                if(string.Equals(Patronymic, other.Patronymic, StringComparison.InvariantCulture))
                {
                    return true;
                }
                else if(string.IsNullOrEmpty(Patronymic) || string.IsNullOrEmpty(other.Patronymic))
                {
                    return true;
                }
            }

            return false;
        }

        public override bool Equals(object obj)
        {
            if (obj is null) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != GetType()) return false;
            return Equals((Person)obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = (FirstName != null ? FirstName.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (Patronymic != null ? Patronymic.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (LastName != null ? LastName.GetHashCode() : 0);
                return hashCode;
            }
        }

        public override string ToString()
        {
            return Patronymic != string.Empty ? $"{LastName} {FirstName} {Patronymic}" : $"{FirstName} {LastName}";
        }
    }
}