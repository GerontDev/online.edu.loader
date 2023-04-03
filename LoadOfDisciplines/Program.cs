using ClosedXML.Excel;

namespace LoadOfDisciplines;

public class Program
{
	public static void Main(string[] args)
	{
		Console.WriteLine("Run");
		using var workbook = new XLWorkbook(@"..\..\..\..\Шаблон дисциплины.xlsx");
		var worksheet = workbook.Worksheets.First();
		//worksheet.Cell("A3").Value = Guid.NewGuid().ToString();
		for (int r = 3; r < ushort.MaxValue; r++)
		{
			var titleCell = worksheet.Cell(r, 2);
			if(!titleCell.Value.IsText || string.IsNullOrEmpty(titleCell.Value.GetText()))
				break;
			worksheet.Cell(r, 1).Value = Guid.NewGuid().ToString();
			worksheet.Cell(r, 3).Value = Guid.NewGuid().ToString();
		}
		workbook.Save();
		Console.WriteLine("Complied");
	}
}