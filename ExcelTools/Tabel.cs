using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;

namespace ZsuTools
{
	/// <summary>
	/// "Табель" table
	/// </summary>
	public class Tabel
	{
		// Allowed state codes (trimmed, case-sensitive using Ordinal comparison)
		// private static readonly HashSet<string> AllowedStates = new HashSet<string>(StringComparer.Ordinal)
		// {
		// 	"+",
		// 	"-",
		// 	"++",
		// 	"вдр",
		// 	"від",
		// 	"лік",
		// 	"СЗЧ",
		// 	""
		// };
		public static List<TabelEntry> ReadTabel(Worksheet ws)
		{
			if (ws == null) return null;

			var result = new List<TabelEntry>();

			// Find the last row (by PIB column)
			Range lastRowRange = ExcelUtils.GetLastSingleNonEmptyCell(ws, 2);
			if (lastRowRange == null) return result;
			var endRow = lastRowRange.Row;

			// Find the last rightmost day number cell
			const String firstDayNumberCell = "C4";
			Range  lastColRange = ws.Range[firstDayNumberCell].End[XlDirection.xlToRight];
			var  endCol = lastColRange.Column;

			// 2. Основний цикл читання з рядка 5
			for (int row = 5; row <= endRow; row++)
			{
				Range cellA = ws.Cells[row, 1];
				if (cellA.MergeCells is bool m && m)
					continue;

				// --- ЛОГІКА ОБРОБКИ РЯДКІВ ВІЙСЬКОВОСЛУЖБОВЦІВ ---
				object valA = cellA.Value2;

				// Якщо необ'єднана і пуста - пропускаємо
				if (valA == null) continue;

				// Якщо непуста - створюємо запис
				var entry = new TabelEntry
				{
					Rank = valA.ToString(),
					FullName = ws.Cells[row, 2].Value2?.ToString() ?? ""
				};

				// Parse states from columns C..endCol
				var (states, isEmptyRanges) = ParseStates(ws, row, 3, endCol);
				entry.States = states;
				entry.ContainsEmptyRanges = isEmptyRanges;

				result.Add(entry);
			}

			return result;
		}

		/// <summary>
		/// Parse consecutive state codes from a worksheet row between startCol and endCol (inclusive).
		/// Column C corresponds to day 1.
		/// </summary>
		private static (List<TabelEntry.StateRange>, bool) ParseStates(Worksheet ws, int row, int startCol, int endCol)
		{
			var list = new List<TabelEntry.StateRange>();
			if (endCol < startCol) return (list, false);

			bool containsEmpty = false;
			string currentState = null;
			int currentStart = 0;
			int statesCount = endCol - startCol + 1;
			//var allowed = AllowedStates;
			for (int col = startCol; col <= endCol; col++)
			{
				int day = col - startCol + 1;
				string s = ws.Cells[row, col].Value2?.ToString() ?? "";
				s = s.Trim();
				// if (!allowed.Contains(s))
				// {
				// 	throw new InvalidOperationException($"Invalid state '{s}' at row {row}, column {col} of Tabel");
				// }
				if(s == "")					
					containsEmpty = true;

				if (currentState == null)
				{
					currentState = s;
					currentStart = day;
				}
				else if (!string.Equals(currentState, s, StringComparison.Ordinal))
				{
					// only add non-empty states to the result list
					if (!string.IsNullOrEmpty(currentState))
					{
						list.Add(new TabelEntry.StateRange { State = currentState, StartDay = currentStart, EndDay = day - 1 });
					}
					currentState = s;
					currentStart = day;
				}
			}
			if (currentState != null)
			{
				if (!string.IsNullOrEmpty(currentState))
				{
					list.Add(new TabelEntry.StateRange { State = currentState, StartDay = currentStart, EndDay = statesCount });
				}
			}
			return (list, containsEmpty);
		}
	}
}
