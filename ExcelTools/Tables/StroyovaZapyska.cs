using System;
using System.Collections.Generic;
using System.Drawing;
using ClosedXML.Excel;
using Microsoft.Office.Interop.Excel;

namespace ZsuTools.Tables
{
    public class StroyovaZapyska
    {
        public IReadOnlyList<Tuple<RankPerson, SZState, Color>> Items => _items;

        public StroyovaZapyska(Worksheet ws)
        {
            _items = SZParser.ParseMilitaryData(ws);
        }

        public StroyovaZapyska(string fileName)
        {
            using (var workbook = new XLWorkbook(fileName))
            {
                var worksheet = workbook.Worksheet(0); 
                
                // Find the last row that contains data
                var lastRow = worksheet.LastRowUsed().RowNumber();
                
                
            }
        }

        private readonly List<Tuple<RankPerson, SZState, Color>> _items;
    }
}