using ClosedXML.Excel;
using Service.Core.Model;

namespace Service.Core
{
	public static class DisciplinesExcel
	{
		public static IReadOnlyList<Discipline> Load(string excelFile)
		{
			using var workbook = new XLWorkbook(excelFile);
			IXLWorksheet? worksheet = workbook.Worksheets.First();
			StructureExcel.CheckHeaderExcel(worksheet, StructureExcel.Disciplines.HeaderColumns);

			List<Discipline> list = new();

			for (int rowIndex = 3; rowIndex < ushort.MaxValue; rowIndex++)
			{
				var titleCell = worksheet.Cell(rowIndex, StructureExcel.Disciplines.Columns.TileColumnNumber);

				if (!titleCell.Value.IsText || string.IsNullOrEmpty(titleCell.Value.GetText()))
					break;

				string title = titleCell.Value.GetText();

				var externalIdCell = worksheet.Cell(rowIndex, StructureExcel.Disciplines.Columns.ExternalIdColumnNumber);
				var idCell = worksheet.Cell(rowIndex, StructureExcel.Disciplines.Columns.IdColumnNumber);

				CheckCellType(rowIndex, externalIdCell, idCell);

				var discipline = new Discipline()
				{
					ExternalId = externalIdCell.Value.GetText(),
					Title = title,
					Id = idCell.Value.GetText(),
				};

				discipline.Check(rowIndex);

				list.Add(discipline);
			}

			return list;
		}

		public static void CheckCellType(int rowIndex, IXLCell? externalIdCell, IXLCell? idCell)
		{
			if (externalIdCell is not null &&  !externalIdCell.Value.IsText)
				throw new Exception(
					$"Excel of Disciplines, row {rowIndex} External Id is {externalIdCell.Value.Type}, should be text");

			if (idCell is not null && !idCell.Value.IsText)
				throw new Exception(
					$"Excel of Disciplines, row {rowIndex} Id is {idCell.Value.Type}, should be text");
		}

		public static void Check(this Discipline discipline, int rowIndex)
		{
			if (string.IsNullOrEmpty(discipline.Title))
				throw new Exception($"Excel file, row {rowIndex} Title is empty");

			if (string.IsNullOrEmpty(discipline.ExternalId))
				throw new Exception($"Excel file, row {rowIndex} External Id is empty");

			if (string.IsNullOrEmpty(discipline.Id))
				throw new Exception($"Excel file, row {rowIndex} Id is empty");
		}
	}
}
