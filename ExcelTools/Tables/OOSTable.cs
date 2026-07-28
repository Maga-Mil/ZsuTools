using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using ZsuTools.Entities;

namespace ZsuTools.Tables
{
    /// <summary>
    /// Таблиця Облік Особового складу ЄЖООС
    /// </summary>
    public class OOSTable
    {
        public IReadOnlyList<Item> Items { get; private set; }
        
        public OOSTable( Worksheet ws)
        {
            if(ws == null) throw new ArgumentNullException(nameof(ws));
            if(!IsOOSTable(ws))  throw new ArgumentException("Не знайдено таблицію ООС");

            var firstRow = 6; //Hardcode
            var lastRow = ExcelUtils.GetLastSingleNonEmptyCell(ws, 1);
            
            if(lastRow == null || lastRow.Row < firstRow)
                throw new InvalidOperationException("Таблиця ООС не має даних");

            var items = new List<Item>();
            for (int i = 6; i < lastRow.Row; i++)
            {
                //Mandatory fields
                String rankValue = ws.Cells[i, 1].Value2?.ToString();
                if( string.IsNullOrEmpty(rankValue) )
                    continue;
                
                String pibValue = ws.Cells[i, 2].Value2?.ToString();
                if( string.IsNullOrEmpty(rankValue) )
                    continue;

                var rank = new Rank(rankValue);
                var pib = new Person(pibValue);

                String rnokpp = ws.Cells[i, 16].Value2?.ToString();
                //Validate RNOKPP?
                
                var item = new Item
                {
                    Rank = rank,
                    Person = pib,
                    RNOKPP = rnokpp
                };

                string positionsString = ws.Cells[i, 3].Value2?.ToString();
                if( !string.IsNullOrEmpty(positionsString) )
                {
                    var positions = positionsString.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var position in positions)
                    {
                        if(int.TryParse(position.Trim(), out int pos))
                        {
                            if(item.Position == null)
                                item.Position = new List<int>();
                            item.Position.Add(pos);
                        }
                    }
                }

                items.Add(item);
                
            }
            
            Items = items;
        }
        
        public static bool IsOOSTable( Worksheet ws )
        {
            return ws.Name == "2. ООС"; 
        }

        public class Item
        {
            public Rank Rank { get; set; }
            public Person Person { get; set; }
            public List<int> Position { get; set; }
            public string RNOKPP { get; set; }
        }
    }
}