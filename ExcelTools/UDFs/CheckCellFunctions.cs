using System;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;

namespace ZsuTools
{
    public class CheckCellFunctions
    {
        [ExcelFunction(
            Name = "IsGreenFont", // Назва функції в Excel (опціонально, за замовчуванням дорівнює імені методу)
            Description = "Перевіряє, чи має текст у комірці зелений колір. Повертає TRUE або FALSE.",
            Category = "ZsuTools", // Функція буде згрупована у цій категорії в Excel
            IsMacroType = true)]
        public static object IsGreenFont(
            [ExcelArgument(Name = "комірка", Description = "Посилання на одну комірку або діапазон (наприклад, A1)",
                AllowReference = true)]
            object cellRef)
        {
            return CheckCellPredicate(cellRef, IsCellGreen);
        }

        [ExcelFunction(Name = "IsBoldFont", 
            Description = "Перевіряє, чи має виділений текст у комірці жирним. Повертає TRUE або FALSE.", 
            Category = "ZsuTools", 
            IsMacroType = true)]
        public static object IsBoldFont([ExcelArgument(Name = "комірка", Description = "Посилання на одну комірку або діапазон (наприклад, A1)",
                AllowReference = true)]
            object cellRef)
        {
            return CheckCellPredicate(cellRef, IsCellBold);
        }

        [ExcelFunction(Name = "IsItalicFont", 
            Description = "Перевіряє, чи має виділений текст у комірці курсивом. Повертає TRUE або FALSE.", 
            Category = "ZsuTools", 
            IsMacroType = true)]
        public static object IsItalicFont([ExcelArgument(Name = "комірка", Description = "Посилання на одну комірку або діапазон (наприклад, A1)",
                AllowReference = true)]
            object cellRef)
        {
            return CheckCellPredicate(cellRef, IsCellItalic);
        }

        [ExcelFunction(Name = "IsStrikeoutFont", 
            Description = "Перевіряє, чи має виділений текст у комірці закресленим. Повертає TRUE або FALSE.", 
            Category = "ZsuTools", 
            IsMacroType = true)]
        public static object IsStrikeoutFont([ExcelArgument(Name = "комірка", Description = "Посилання на одну комірку або діапазон (наприклад, A1)",
                AllowReference = true)]
            object cellRef)
        {
            return CheckCellPredicate(cellRef, IsCellStrikeout);
        }

        private static object CheckCellPredicate(object cellRef, Func<Range, bool> predicate)
        {
            if (!(cellRef is ExcelReference))
                return ExcelError.ExcelErrorRef;
            
            try
            {
                var app = (Application)ExcelDnaUtil.Application;
            
                // Отримуємо весь діапазон, який передав користувач
                var address = (string)XlCall.Excel(XlCall.xlfReftext, cellRef, true);
                var inputRange = app.Range[address];
        
                var rows = inputRange.Rows.Count;
                var cols = inputRange.Columns.Count;
        
                // Якщо це лише одна комірка, повертаємо одне булеве значення
                if (rows == 1 && cols == 1)
                {
                    return predicate(inputRange);
                }
        
                // Якщо це діапазон, створюємо двовимірний масив для результату
                object[,] resultMatrix = new object[rows, cols];
        
                for (var r = 1; r <= rows; r++)
                {
                    for (var c = 1; c <= cols; c++)
                    {
                        var cell = (Range)inputRange.Cells[r, c];
                        resultMatrix[r - 1, c - 1] = predicate(cell);
                    }
                }
        
                return resultMatrix; 
            }
            catch (Exception)
            {
                return ExcelError.ExcelErrorValue;
            }
        }

        // Допоміжний метод для перевірки однієї клітинки
        private static bool IsCellGreen(Range cell)
        {
            int rgb = Convert.ToInt32(cell.Font.Color);
            var r = rgb & 0xFF;
            var g = (rgb >> 8) & 0xFF;
            var b = (rgb >> 16) & 0xFF;

            return (g > r && g > b && g > 50);
        }
        
        private static bool IsCellBold(Range cell)
        {
            return (bool)cell.Font.Bold;
        }
        
        private static bool IsCellItalic(Range cell)
        {
            return (bool)cell.Font.Italic;
        }

        private static bool IsCellStrikeout(Range cell)
        {
            return (bool)cell.Font.Strikethrough;
        }
    }
}