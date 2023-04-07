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

			for (int row = 3; row < ushort.MaxValue; row++)
			{
				var titleCell = worksheet.Cell(row, StructureExcel.EducationalPrograms.Columns.TileColumnNumber);

				if (!titleCell.Value.IsText || string.IsNullOrEmpty(titleCell.Value.GetText()))
					break;

				string title = titleCell.Value.GetText();

				
				var externalIdCell = worksheet.Cell(row, StructureExcel.EducationalPrograms.Columns.ExternalIdColumnNumber);
				var directionCell = worksheet.Cell(row, StructureExcel.EducationalPrograms.Columns.DirectionColumnNumber);
				var codeDirectionCell = worksheet.Cell(row, StructureExcel.EducationalPrograms.Columns.CodeDirectionColumnNumber);
				var startYearCell = worksheet.Cell(row, StructureExcel.EducationalPrograms.Columns.StartYearColumnNumber);
				var endYearCall = worksheet.Cell(row, StructureExcel.EducationalPrograms.Columns.EndYearColumnNumber);
				var idCell = worksheet.Cell(row, StructureExcel.EducationalPrograms.Columns.IdColumnNumber);

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

				educationalProgram.Check(row);

				list.Add(educationalProgram);
			}

			return list;
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
