using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using ClosedXML.Excel;
using Microsoft.Office.Interop.Excel;

namespace ZsuTools.GenerateTabel
{
    public class СтройоваЗаписка
    {
        public DateTime Date { get; private set; }
        
        public List<Tuple<RankPerson, SZState, Color>> MilitaryData { get; private set; }
        
        public СтройоваЗаписка(Workbook стройоваЗапискаWb)
        {
            //Check for nonempty workbook
            if (стройоваЗапискаWb == null || стройоваЗапискаWb.Worksheets.Count == 0)
            {
                throw new ArgumentException("Invalid workbook provided.");
            }
            
            //Open first worksheet, its always Стройова Записка
            Worksheet worksheet = стройоваЗапискаWb.Worksheets[1];

            Date = SZParser.ParseDate(worksheet);
            MilitaryData = SZParser.ParseMilitaryData(worksheet);
        }

        /// <summary>
        /// Створює Стройову Записку з Excel файлу, використовуючи ClosedXML (без Excel Interop).
        /// </summary>
        public СтройоваЗаписка(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                throw new ArgumentException("Invalid file path provided.", nameof(filePath));
            }

            using (var workbook = new XLWorkbook(filePath))
            {
                Load(workbook);
            }
        }

        /// <summary>
        /// Створює Стройову Записку з Stream Excel файлу, використовуючи ClosedXML (без Excel Interop).
        /// </summary>
        public СтройоваЗаписка(Stream stream)
        {
            if (stream == null)
            {
                throw new ArgumentException("Invalid stream provided.", nameof(stream));
            }

            using (var workbook = new XLWorkbook(stream))
            {
                Load(workbook);
            }
        }

        private void Load(XLWorkbook workbook)
        {
            if (workbook.Worksheets.Count == 0)
            {
                throw new ArgumentException("Invalid workbook provided.");
            }

            // First worksheet is always Стройова Записка (ClosedXML worksheets are 1-based)
            var worksheet = workbook.Worksheet(1);
            Date = SZParser.ParseDate(worksheet);
            MilitaryData = SZParser.ParseMilitaryData(worksheet);
        }
    }
}