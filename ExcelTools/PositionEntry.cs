using System.Diagnostics;

namespace ZsuTools
{
	[DebuggerDisplay("{Index}-{PersonName}")]
	public class PositionEntry
	{
		public int Index { get; set; }
		public string SubUnit { get; set; }
		/// <summary>
		/// Посада
		/// </summary>
		public string Name { get; set; }
		/// <summary>
		/// ШПК
		/// </summary>
		public string AuthorizedRank { get; set; }
		/// <summary>
		/// ВОС
		/// </summary>
		public string MOS { get; set; }
		/// <summary>
		/// Звання
		/// </summary>
		public string PersonRank { get; set; }
		/// <summary>
		/// ПІБ
		/// </summary>
		public string PersonName { get; set; }
	}
}
