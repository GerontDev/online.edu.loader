using ClosedXML.Excel;

namespace Service.Core;

public static class StructureExcel
{
	public static class Disciplines
	{
		public static string[] HeaderColumns = { "external_id", "title", "ID" };
		public static class Columns
		{
			public const int ExternalIdColumnNumber = 1;
			public const int TileColumnNumber = 2;
			public const int IdColumnNumber = 3;
		}
	}

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

	public static class Student
	{
		public static string[] HeaderColumns = { "surname", "name", "middle_name", "snils", "inn", "email", "external_id", "study_year", "ID1", "title1", "study_plan","ID2" };
		public static class ColumnNumber
		{
			public const int Surname = 1;
			public const int Name = 2;
			public const int MiddleName = 3;
			public const int Snils = 4;
			public const int Inn = 5;
			public const int Email = 6;
			public const int ExternalId = 7;
			public const int StudyYear = 8;
			public const int Id1 = 9;
			public const int Title1 = 10;
			public const int StudyPlan = 11;
			public const int Id2 = 12;
		}
	}

	public static bool IsInvalidExcelFile(IXLWorksheet worksheet, string[] columns)
	{
		for (int c = 1; c <= columns.Length; c++)
		{
			if (worksheet.Cell(1, c).Value.GetText() != columns[c - 1])
				return true;
		}

		return false;
	}

	public static void CheckHeaderExcel(IXLWorksheet worksheet, string[] columns)
	{
		if (IsInvalidExcelFile(worksheet, columns))
			throw new Exception($"File is invalid. Columns is need [{string.Join(", ", columns)}] ");
	}
}

