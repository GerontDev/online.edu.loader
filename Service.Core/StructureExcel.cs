using ClosedXML.Excel;

namespace Service.Core;

public static class StructureExcel
{
	
	public static class EducationalPrograms
	{
		public static string[] HeaderColumns = { "external_id", "title", "direction", "code_direction", "start_year", "end_year", "ID" };
		public static class Columns
		{
			public const int ExternalIdColumnNumber = 1;
			public const int TileColumnNumber = 2;
			public const int DirectionColumnNumber = 3;
			public const int CodeDirectionColumnNumber = 4;
			public const int StartYearColumnNumber = 5;
			public const int EndYearColumnNumber = 6;
			public const int IdColumnNumber = 7;
		}
	}

	public static class StudyPlans
	{
		public static string[] HeaderColumns = { "external_id", "title", "direction", "code_direction", "start_year", "end_year", "education_form", "educational_program", "ID" };
		public static class Columns
		{
			public const int ExternalIdColumnNumber = 1;
			public const int TileColumnNumber = 2;
			public const int DirectionColumnNumber = 3;
			public const int CodeDirectionColumnNumber = 4;
			public const int StartYearColumnNumber = 5;
			public const int EndYearColumnNumber = 6;
			public const int EducationFormColumnNumber = 7;
			public const int EducationalProgramColumnNumber = 8;
			public const int IdColumnNumber = 9;
		}
	}

	public static class LinkDisciplinesAndStudyPlans
	{
		public static string[] HeaderColumns = { "title1", "title2", "study_plan", "discipline", "semester", "ID"};
		public static class Columns
		{
			public const int Title1ColumnNumber = 1;
			public const int Title2ColumnNumber = 2;
			public const int StudyPlanColumnNumber = 3;
			public const int DisciplineColumnNumber = 4;
			public const int SemesterColumnNumber = 5;
			public const int IDColumnNumber = 6;
		}
	}

	public static bool IsInvalidExcelFile(IXLWorksheet worksheet, string[] columns)
	{
		for (int c = 1; c <= columns.Length; c++)
		{
			if (worksheet.Cell(c, 1).Value.GetText() != columns[c - 1])
				return false;
		}

		return true;
	}

	public static void CheckHeaderExcel(IXLWorksheet worksheet, string[] columns)
	{
		if (IsInvalidExcelFile(worksheet, columns))
			throw new Exception($"File is invalid. Columns is need [{string.Join(", ", columns)}] ");
	}
}

