using System;
using System.Drawing;
using ClosedXML.Excel;
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
        
        /// <summary>
        /// Колір шрифту комірки через ClosedXML (аналог <see cref="GetCellFontColor(Range)"/> для COM).
        /// </summary>
        public static Color GetCellFontColor(IXLCell cell)
        {
            try
            {
                var xlColor = cell.Style.Font.FontColor;
                if (xlColor == null || !xlColor.HasValue) return Color.Black;

                if (xlColor.ColorType == XLColorType.Theme)
                {
                    // XLColor.Color не працює для кольорів з теми - розв'язуємо через тему книги
                    var baseColor = cell.Worksheet.Workbook.Theme.ResolveThemeColor(xlColor.ThemeColor).Color;
                    return ApplyTint(baseColor, xlColor.ThemeTint);
                }

                return xlColor.Color;
            }
            catch
            {
                return Color.Black;
            }
        }

        /// <summary>
        /// Застосовує tint так само, як Excel: зміна яскравості (L) у просторі HSL.
        /// tint &lt; 0 - темніше, tint &gt; 0 - світліше.
        /// </summary>
        private static Color ApplyTint(Color color, double tint)
        {
            if (Math.Abs(tint) < 1e-9) return color;

            double h = color.GetHue();
            double s = color.GetSaturation();
            double l = color.GetBrightness();

            l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;

            return FromHsl(h, s, l);
        }

        private static Color FromHsl(double h, double s, double l)
        {
            double c = (1 - Math.Abs(2 * l - 1)) * s;
            double x = c * (1 - Math.Abs(h / 60 % 2 - 1));
            double m = l - c / 2;

            double r, g, b;
            if (h < 60) { r = c; g = x; b = 0; }
            else if (h < 120) { r = x; g = c; b = 0; }
            else if (h < 180) { r = 0; g = c; b = x; }
            else if (h < 240) { r = 0; g = x; b = c; }
            else if (h < 300) { r = x; g = 0; b = c; }
            else { r = c; g = 0; b = x; }

            return Color.FromArgb(
                (int)Math.Round((r + m) * 255),
                (int)Math.Round((g + m) * 255),
                (int)Math.Round((b + m) * 255));
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