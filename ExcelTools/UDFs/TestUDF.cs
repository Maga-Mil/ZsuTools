using ExcelDna.Integration;

namespace ZsuTools
{
	public static class TestUDF
	{
		[ExcelFunction(Description = "Отримує адресу діапазону")]                        
		public static int TestGetDateRanges2(object[,] input)
		{
			return input.GetLength(0) + input.GetLength(1);
		}
	}
}
