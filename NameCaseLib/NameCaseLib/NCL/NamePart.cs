namespace NameCaseLib.NCL
{
    /// <summary>
    ///     Перечисление для идентификации типа слова
    /// </summary>
    public enum NamePart
    {
        /// <summary>
        ///     Слово не идентифицировано
        /// </summary>
        Null = 0,

        /// <summary>
        ///     Имя
        /// </summary>
        Name = 1,

        /// <summary>
        ///     Фамилия
        /// </summary>
        Surname = 2,

        /// <summary>
        ///     Отчество
        /// </summary>
        FatherName = 3
    }
}