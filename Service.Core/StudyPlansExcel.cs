using ClosedXML.Excel;
using Service.Core.Model;

namespace Service.Core
{
	public static class StudyPlansExcel
	{
		public static IReadOnlyList<StudyPlan> Load(string excelFile)
		{
			using var workbook = new XLWorkbook(excelFile);
			IXLWorksheet? worksheet = workbook.Worksheets.First();
			StructureExcel.CheckHeaderExcel(worksheet, StructureExcel.StudyPlans.HeaderColumns);

			List<StudyPlan> list = new();

			for (int rowIndex = 3; rowIndex < ushort.MaxValue; rowIndex++)
			{
				var titleCell = worksheet.Cell(rowIndex, StructureExcel.StudyPlans.Columns.TileColumnNumber);

				if (!titleCell.Value.IsText || string.IsNullOrEmpty(titleCell.Value.GetText()))
					break;

				string title = titleCell.Value.GetText();

				var externalIdCell = worksheet.Cell(rowIndex, StructureExcel.StudyPlans.Columns.ExternalIdColumnNumber);
				var directionCell = worksheet.Cell(rowIndex, StructureExcel.StudyPlans.Columns.DirectionColumnNumber);
				var codeDirectionCell = worksheet.Cell(rowIndex, StructureExcel.StudyPlans.Columns.CodeDirectionColumnNumber);
				var startYearCell = worksheet.Cell(rowIndex, StructureExcel.StudyPlans.Columns.StartYearColumnNumber);
				var endYearCall = worksheet.Cell(rowIndex, StructureExcel.StudyPlans.Columns.EndYearColumnNumber);
				var educationFormCell = worksheet.Cell(rowIndex, StructureExcel.StudyPlans.Columns.EducationFormColumnNumber);
				var educationalProgramCell = worksheet.Cell(rowIndex, StructureExcel.StudyPlans.Columns.EducationalProgramColumnNumber);
				var idCell = worksheet.Cell(rowIndex, StructureExcel.StudyPlans.Columns.IdColumnNumber);

				CheckCellType(rowIndex, externalIdCell, directionCell, codeDirectionCell, startYearCell, endYearCall, educationFormCell, educationalProgramCell, idCell);

				var studyPlan = new StudyPlan()
				{
					ExternalId = externalIdCell.Value.GetText(),
					Title = title,
					Direction = directionCell.Value.GetText(),
					CodeDirection = codeDirectionCell.Value.GetText(),
					EndYear = (int)startYearCell.Value.GetNumber(),
					StartYear = (int)endYearCall.Value.GetNumber(),
					EducationalProgram = educationalProgramCell.Value.GetText(),
					EducationForm = educationFormCell.Value.GetText(),
					Id = idCell.Value.GetText(),
				};

				studyPlan.Check(rowIndex);

				list.Add(studyPlan);
			}

			return list;
		}

		public static void CheckCellType(int rowIndex, IXLCell? externalIdCell, IXLCell directionCell,
			IXLCell codeDirectionCell, IXLCell startYearCell, IXLCell endYearCall, IXLCell educationFormCell, IXLCell educationalProgramCell, IXLCell? idCell)
		{
			if (externalIdCell is not null &&  !externalIdCell.Value.IsText)
				throw new Exception(
					$"Excel of StudyPlans, row {rowIndex} External Id is {externalIdCell.Value.Type}, should be text");
			if (!directionCell.Value.IsText)
				throw new Exception(
					$"Excel of StudyPlans, row {rowIndex} Direction is {directionCell.Value.Type}, should be text");
			if (!codeDirectionCell.Value.IsText)
				throw new Exception(
					$"Excel of StudyPlans, row {rowIndex} Code Direction is {codeDirectionCell.Value.Type}, should be text");
			if (!startYearCell.Value.IsNumber)
				throw new Exception(
					$"Excel of StudyPlans, row {rowIndex} Start Year is {startYearCell.Value.Type}, should be number");
			if (!endYearCall.Value.IsNumber)
				throw new Exception(
					$"Excel of StudyPlans, row {rowIndex} End Year is {endYearCall.Value.Type}, should be number");

			if (!educationFormCell.Value.IsText)
				throw new Exception(
					$"Excel of StudyPlans, row {rowIndex} Education Form is {educationFormCell.Value.Type}, should be text");

			if (!educationalProgramCell.Value.IsText)
				throw new Exception(
					$"Excel of StudyPlans, row {rowIndex} Educational Program is {educationalProgramCell.Value.Type}, should be text");

			if (idCell is not null && !idCell.Value.IsText)
				throw new Exception(
					$"Excel of StudyPlans, row {rowIndex} Id is {idCell.Value.Type}, should be text");
		}

		public static void Check(this StudyPlan studyPlan, int rowIndex)
		{
			if (string.IsNullOrEmpty(studyPlan.Direction))
				throw new Exception($"Excel file, row {rowIndex} direction is empty");

			if (string.IsNullOrEmpty(studyPlan.CodeDirection))
				throw new Exception($"Excel file, row {rowIndex} Code Direction is empty");

			if (studyPlan.StartYear < 1900)
				throw new Exception($"Excel file, row {rowIndex} Start Year is invalid");

			if (studyPlan.EndYear < 1900)
				throw new Exception($"Excel file, row {rowIndex} End Year is invalid");

			if (string.IsNullOrEmpty(studyPlan.EducationForm))
				throw new Exception($"Excel file, row {rowIndex} Education Form is empty");

			if (string.IsNullOrEmpty(studyPlan.EducationalProgram))
				throw new Exception($"Excel file, row {rowIndex} Educational Program is empty");
		}
	}
}
