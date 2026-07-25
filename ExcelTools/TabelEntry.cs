using System.Collections.Generic;

namespace ZsuTools
{
	/// <summary>
	/// One row from Tabel table
	/// </summary>
	public class TabelEntry
	{
		/// <summary>
		/// Person rank
		/// </summary>
		public string Rank { get; set; }

		/// <summary>
		/// Person PIB
		/// </summary>
		public string FullName { get; set; }

		/// <summary>
		/// List of tabel ranges
		/// </summary>
		public List<StateRange> States { get; set; } = new List<StateRange>();

		/// <summary>
		/// If true - some cells of the Tabel row contains empty value
		/// </summary>
		public bool ContainsEmptyRanges { get; set; }

		public class StateRange
		{
			public string State { get; set; }
			public int StartDay { get; set; }
			public int EndDay { get; set; }
		}
	}
}

