using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;

namespace ZsuTools
{
	/// <summary>
	/// ШПО table
	/// </summary>
	public class Positions
	{
		/// <summary>
		/// Read positions table from the first worksheet whose name contains "ШПО".
		/// Starts at A6 and for each row:
		/// - If cell A is merged: remember the merged value as SubUnit.
		/// - If cell A is not merged: build a PositionEntry from the row:
		///     A - Index, B - Name, C - AuthorizedRank, D - MOS, F - Rank, G - FullName
		/// Processing stops when a row has an empty (non-merged) column A.
		/// Returns an empty list when workbook or sheet is not available.
		/// </summary>
		public static List<PositionEntry> ReadPositionsTable( Worksheet ws )
		{
			var result = new List<PositionEntry>();

			if ( ws == null )
				return result;

			dynamic excelApp = null;
			Workbook workbook = null;

			try
			{
				excelApp = ExcelDnaUtil.Application;
				if (excelApp == null) return result;

				// Prefer ActiveWorkbook, fall back to first open workbook
				workbook = (Workbook)(excelApp.ActiveWorkbook ?? (excelApp.Workbooks.Count > 0 ? excelApp.Workbooks[1] : null));
				if (workbook == null) return result;

				int startRow = 6;
				string currentSubUnit = string.Empty;

				for (int r = startRow; ; r++)
				{
					Range cellA = null;
					Range cellB = null;
					Range cellC = null;
					Range cellD = null;
					Range cellF = null;
					Range cellG = null;

					try
					{
						cellA = ws.Cells[r, 1] as Range;
						cellB = ws.Cells[r, 2] as Range;
						cellC = ws.Cells[r, 3] as Range;
						cellD = ws.Cells[r, 4] as Range;
						cellF = ws.Cells[r, 6] as Range;
						cellG = ws.Cells[r, 7] as Range;

						// If we couldn't acquire cellA (unexpected), stop
						if (cellA == null)
							break;

						// Check whether cell A is part of merged area (simplified)
						bool isMerged = false;
						try
						{
							// MergeCells typically returns a bool; evaluate safely
							isMerged = cellA.MergeCells is bool b && b;
						}
						catch
						{
							isMerged = false;
						}

						// Read raw value of A (if available)
						object aValue = null;
						try { aValue = cellA.Value2; } catch { aValue = null; }

						// If not merged and truly empty -> stop processing
						if (!isMerged && aValue == null)
						{
							break;
						}

						if (isMerged)
						{
							// merged row acts as SubUnit header
							Range mergeArea = null;
							try
							{
								mergeArea = cellA.MergeArea;
								object mv = null;
								try { mv = mergeArea.Cells[1, 1].Value2; } catch { mv = null; }
								var subUnitValue = mv != null ? mv.ToString().Trim() : string.Empty;
								currentSubUnit = ToSentenceCase(subUnitValue);
							}
							finally
							{
								if (mergeArea != null) { Marshal.ReleaseComObject(mergeArea); mergeArea = null; }
							}

							// continue to next row (merged rows are headers)
							continue;
						}
						else
						{
							// Non-merged row — read a PositionEntry
							var entry = new PositionEntry();

							// Index (col A) - try numeric then fallback to parse
							if (aValue is double da)
								entry.Index = Convert.ToInt32(da);
							else if (aValue is int ia)
								entry.Index = ia;
							else if (aValue != null && int.TryParse(aValue.ToString(), out int pi))
								entry.Index = pi;
							else
								entry.Index = 0;

							entry.SubUnit = currentSubUnit ?? string.Empty;
							entry.Name = cellB?.Value2 != null ? cellB.Value2.ToString().Trim() : string.Empty;
							entry.AuthorizedRank = cellC?.Value2 != null ? cellC.Value2.ToString().Trim() : string.Empty;
							entry.MOS = cellD?.Value2 != null ? cellD.Value2.ToString().Trim() : string.Empty;
							entry.PersonRank = cellF?.Value2 != null ? cellF.Value2.ToString().Trim() : string.Empty;
							entry.PersonName = cellG?.Value2 != null ? cellG.Value2.ToString().Trim() : string.Empty;

							// Skip rows that seem completely empty
							if (entry.Index == 0 && string.IsNullOrEmpty(entry.Name) && string.IsNullOrEmpty(entry.PersonName))
							{
								// skip this row but continue scanning (A was non-empty)
							}
							else
							{
								result.Add(entry);
							}
						}
					}
					finally
					{
						if (cellA != null) { Marshal.ReleaseComObject(cellA); cellA = null; }
						if (cellB != null) { Marshal.ReleaseComObject(cellB); cellB = null; }
						if (cellC != null) { Marshal.ReleaseComObject(cellC); cellC = null; }
						if (cellD != null) { Marshal.ReleaseComObject(cellD); cellD = null; }
						if (cellF != null) { Marshal.ReleaseComObject(cellF); cellF = null; }
						if (cellG != null) { Marshal.ReleaseComObject(cellG); cellG = null; }
					}
				}
			}
			finally
			{
				if (ws != null) { Marshal.ReleaseComObject(ws); ws = null; }
				if (workbook != null) { Marshal.ReleaseComObject(workbook); workbook = null; }
				// ExcelDnaUtil.Application is owned by Excel — do not release.
			}

			return result;
		}
		/// <summary>
		/// Convert a string to sentence case (first letter uppercase, rest lowercase).
		/// </summary>
		private static string ToSentenceCase(string text)
		{
			if (string.IsNullOrEmpty(text))
				return text;
			
			return char.ToUpper(text[0]) + text.Substring(1).ToLower();
		}
	}
}
