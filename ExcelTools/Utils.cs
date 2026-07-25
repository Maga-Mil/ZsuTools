using System;
using System.Drawing;
using Microsoft.Office.Interop.Excel;

namespace ZsuTools
{
    public static class Utils
    {
        /// <summary>
        /// Because we use .Net 4.8 instead if .Net Core
        /// </summary>
        public static bool Contains(this String str, String substring, 
            StringComparison comp)
        {                            
            if (substring == null)
                throw new ArgumentNullException("substring", 
                    "substring cannot be null.");
            else if (!Enum.IsDefined(typeof(StringComparison), comp))
                throw new ArgumentException("comp is not a member of StringComparison",
                    "comp");

            return str.IndexOf(substring, comp) >= 0;                      
        }
        
        public static string FirstLetterUppercase(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;

            return char.ToUpperInvariant(input[0]) + input.Substring(1).ToLowerInvariant();
        }
        
        public static Color GetCellFontColor(Range cell)
        {
            if (cell == null || cell.Font.Color == null) 
                return Color.Black;

            // 1. Отримуємо значення кольору як об'єкт
            object rawColor = cell.Font.Color;

            // 2. Перевіряємо, чи це не константа "Автоматично"
            // xlColorIndexAutomatic дорівнює -4105
            if (rawColor is int colorInt && colorInt == -4105)
            {
                // Якщо колір автоматичний, повертаємо дефолтний чорний
                return Color.Black; 
            }

            // 3. Якщо колір фіксований (користувач вибрав його вручну),
            // безпечно конвертуємо через OLE
            try
            {
                int oleColor = Convert.ToInt32(rawColor);
                return ColorTranslator.FromOle(oleColor);
            }
            catch
            {
                return Color.Black; // На випадок непередбачуваних помилок COM
            }
        }
        
        public static Color GetCellFillColor(Range cell)
        {
            if (cell == null || cell.Interior.Color == null) 
                return Color.Transparent;

            if (cell.Interior.ColorIndex == -4142)
            {
                // Якщо колір автоматичний, повертаємо дефолтний чорний
                return Color.Transparent; 
            }
            
            object rawColor = cell.Interior.Color;
            if (rawColor is int intColor)
            {
                return ColorTranslator.FromOle(intColor);
            }
            else
            {
                return Color.Transparent;
            }
        }
        
        
    }
}