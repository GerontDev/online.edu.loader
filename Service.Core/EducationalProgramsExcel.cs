using ClosedXML.Excel;
using Service.Core.Model;

namespace Service.Core
{
	public static class EducationalProgramsExcel
	{
		public static IReadOnlyList<EducationalProgram> Load(string excelFile)
		{
			using var workbook = new XLWorkbook(excelFile);
			IXLWorksheet? worksheet = workbook.Worksheets.First();
			StructureExcel.CheckHeaderExcel(worksheet, StructureExcel.EducationalPrograms.HeaderColumns);

			List<EducationalProgram> list = new();

			for (int rowIndex = 3; rowIndex < ushort.MaxValue; rowIndex++)
			{
				var titleCell = worksheet.Cell(rowIndex, StructureExcel.EducationalPrograms.Columns.TileColumnNumber);

				if (!titleCell.Value.IsText || string.IsNullOrEmpty(titleCell.Value.GetText()))
					break;

				string title = titleCell.Value.GetText();

				var externalIdCell = worksheet.Cell(rowIndex, StructureExcel.EducationalPrograms.Columns.ExternalIdColumnNumber);
				var directionCell = worksheet.Cell(rowIndex, StructureExcel.EducationalPrograms.Columns.DirectionColumnNumber);
				var codeDirectionCell = worksheet.Cell(rowIndex, StructureExcel.EducationalPrograms.Columns.CodeDirectionColumnNumber);
				var startYearCell = worksheet.Cell(rowIndex, StructureExcel.EducationalPrograms.Columns.StartYearColumnNumber);
				var endYearCall = worksheet.Cell(rowIndex, StructureExcel.EducationalPrograms.Columns.EndYearColumnNumber);
				var idCell = worksheet.Cell(rowIndex, StructureExcel.EducationalPrograms.Columns.IdColumnNumber);

				CheckCellType(rowIndex, externalIdCell, directionCell, codeDirectionCell, startYearCell, endYearCall, idCell);

				var educationalProgram = new EducationalProgram()
				{
					ExternalId = externalIdCell.Value.GetText(),
					Title = title,
					Direction = directionCell.Value.GetText(),
					CodeDirection = codeDirectionCell.Value.GetText(),
					EndYear = (int)startYearCell.Value.GetNumber(),
					StartYear = (int)endYearCall.Value.GetNumber(),
					Id = idCell.Value.GetText(),
				};

				educationalProgram.Check(rowIndex);

				list.Add(educationalProgram);
			}

			return list;
		}

		public static void CheckCellType(int rowIndex, IXLCell? externalIdCell, IXLCell directionCell,
			IXLCell codeDirectionCell, IXLCell startYearCell, IXLCell endYearCall, IXLCell? idCell)
		{
			if (externalIdCell is not null &&  !externalIdCell.Value.IsText)
				throw new Exception(
					$"Excel of EducationalPrograms, row {rowIndex} External Id is {externalIdCell.Value.Type}, should be text");
			if (!directionCell.Value.IsText)
				throw new Exception(
					$"Excel of EducationalPrograms, row {rowIndex} Direction is {directionCell.Value.Type}, should be text");
			if (!codeDirectionCell.Value.IsText)
				throw new Exception(
					$"Excel of EducationalPrograms, row {rowIndex} Code Direction is {codeDirectionCell.Value.Type}, should be text");
			if (!startYearCell.Value.IsNumber)
				throw new Exception(
					$"Excel of EducationalPrograms, row {rowIndex} Start Year is {startYearCell.Value.Type}, should be number");
			if (!endYearCall.Value.IsNumber)
				throw new Exception(
					$"Excel of EducationalPrograms, row {rowIndex} End Year is {endYearCall.Value.Type}, should be number");
			if (idCell is not null && !idCell.Value.IsText)
				throw new Exception(
					$"Excel of EducationalPrograms, row {rowIndex} Id is {idCell.Value.Type}, should be text");
		}

		public static void Check(this EducationalProgram educationalProgram, int rowIndex)
		{
			if (string.IsNullOrEmpty(educationalProgram.Direction))
				throw new Exception($"Excel file, column {rowIndex} direction is empty");

			if (string.IsNullOrEmpty(educationalProgram.CodeDirection))
				throw new Exception($"Excel file, column {rowIndex} Code Direction is empty");

			if (educationalProgram.StartYear < 1900)
				throw new Exception($"Excel file, column {rowIndex} Start Year is invalid");

			if (educationalProgram.EndYear < 1900)
				throw new Exception($"Excel file, column {rowIndex} End Year is invalid");
		}

	}
}
