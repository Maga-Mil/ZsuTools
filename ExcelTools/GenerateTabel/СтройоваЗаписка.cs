using System;
using System.Collections.Generic;
using System.Drawing;
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

            var testResult = SZParser.ParseMilitaryData(worksheet);

            MilitaryData = testResult;
        }
    }
}