namespace NameCaseLib.NCL
{
    /// <summary>
    ///     Класс который содержит тип данных для определения пола человека
    /// </summary>
    public class GenderProbability
    {
        /// <summary>
        ///     Создать новый объект с указанием вероятности принадлежности пола мужчине или женщине
        /// </summary>
        /// <param name="man">Вероятность мужского пола</param>
        /// <param name="woman">Вероятность женского пола</param>
        public GenderProbability(float man, float woman)
        {
            Man = man;
            Woman = woman;
        }

        /// <summary>
        ///     Создание пустого объекта для дальнейшего накопления вероятностей
        /// </summary>
        public GenderProbability()
            : this(0, 0) { }

        /// <summary>
        ///     Получить/Указать вероятность мужского пола
        /// </summary>
        public float Man { get; set; }

        /// <summary>
        ///     Получить/Указать вероятность женского пола
        /// </summary>
        public float Woman { get; set; }

        /// <summary>
        ///     Просуммировать две вероятности
        /// </summary>
        /// <param name="number">Первая вероятность</param>
        /// <param name="add">Вторая вероятность</param>
        /// <returns>Сумма вероятностей</returns>
        public static GenderProbability operator +(GenderProbability number, GenderProbability add)
        {
            var result = new GenderProbability(0, 0)
            {
                Man = number.Man + add.Man,
                Woman = number.Woman + add.Woman
            };

            return result;
        }
    }
}