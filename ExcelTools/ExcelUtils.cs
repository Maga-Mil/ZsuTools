using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace ZsuTools
{
	public static class ExcelUtils
	{
		/// <summary>
		/// Find last not merged not empty cell in the given column. Returns null if no such cell is found.
		/// </summary>
		/// <param name="ws"></param>
		/// <param name="columnNumber"></param>
		/// <returns></returns>
		public static Range GetLastSingleNonEmptyCell(Worksheet ws, int columnNumber)
		{
			if (ws == null) return null;

			// 1. Start from the last possible cell in the worksheet for the given column
			Range lastCellInColumn = ws.Cells[ws.Rows.Count, columnNumber];

			// 2. Jump up to the first cell with content (analogous to Ctrl+Up)
			Range currentCell = lastCellInColumn.End[XlDirection.xlUp];

			// 3. Check the found cell and move upward if necessary
			while (currentCell.Row >= 1)
			{
				// Check for merged cells via dynamic MergeCells property
				bool isMerged = false;
				object mergeValue = currentCell.MergeCells;
				if (mergeValue is bool b) isMerged = b;

				// Success condition: NOT merged and NOT empty
				if (!isMerged && currentCell.Value2 != null)
				{
					return currentCell;
				}

				// If we landed in a merged area — jump above it
				if (isMerged)
				{
					int topRowOfArea = currentCell.MergeArea.Row;
					if (topRowOfArea <= 1) break;
					currentCell = ws.Cells[topRowOfArea - 1, columnNumber];
				}
				else
				{
					// If the cell is simply empty — move one row up
					currentCell = currentCell.Offset[-1, 0];
				}

				// Additional check for possible gaps in data
				if (currentCell.Value2 == null && currentCell.Row > 1)
				{
					currentCell = currentCell.End[XlDirection.xlUp];
				}
				else if (currentCell.Row == 1 && (currentCell.Value2 == null || isMerged))
				{
					break; // Reached top and found nothing usable
				}
			}

			return null;
		}


		/// <summary>
		/// Find first worksheet in workbook whose name contains the given substring (case-insensitive).
		/// Releases COM objects for non-matching sheets. Returns a Worksheet instance (caller must release it).
		/// </summary>
		public static Worksheet FindWorksheetByNameContains(Workbook wb, string namePart)
		{
			if (wb == null || string.IsNullOrEmpty(namePart)) return null;

			Sheets sheets = null;
			Worksheet candidateWs = null;
			try
			{
				sheets = wb.Worksheets;
				foreach (Worksheet s in sheets)
				{
					string name = string.Empty;
					try { name = (s.Name ?? string.Empty).ToString(); } catch { name = string.Empty; }

					if (name.IndexOf(namePart, StringComparison.InvariantCultureIgnoreCase) >= 0)
					{
						// return the found sheet (do not release it here — caller will release)
						candidateWs = s;
						return candidateWs;
					}
					else
					{
						// release this non-matching sheet COM object
						try { Marshal.ReleaseComObject(s); } catch { }
					}
				}
			}
			finally
			{
				if (sheets != null) { Marshal.ReleaseComObject(sheets); sheets = null; }
			}

			return null;
		}

	}
}

