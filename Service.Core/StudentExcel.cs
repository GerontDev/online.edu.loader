using ClosedXML.Excel;
using Service.Core.Model;

namespace Service.Core
{
	public static class StudentExcel
	{
		public static IReadOnlyList<Student> Load(string excelFile)
		{
			using var workbook = new XLWorkbook(excelFile);
			IXLWorksheet? worksheet = workbook.Worksheets.First();
			StructureExcel.CheckHeaderExcel(worksheet, StructureExcel.Student.HeaderColumns);
			ArgumentNullException.ThrowIfNull(worksheet);
			return GetStudents(worksheet).ToList();
		}

		private static IEnumerable<Student> GetStudents(IXLWorksheet worksheet)
		{
			for (int rowIndex = 3; rowIndex < ushort.MaxValue; rowIndex++)
			{
				TupleCells? tuple = worksheet.GetTupleCells(rowIndex);
				if (tuple is null)
					continue;

				tuple.CheckCellType(rowIndex);

				var student = tuple.ToStudent();

				student.Check(rowIndex);
				yield return student;
			}
		}

		public static TupleCells? GetTupleCells(this IXLWorksheet worksheet, int rowIndex)
		{
			var surnameCell = worksheet.Cell(rowIndex, StructureExcel.Student.ColumnNumber.Surname);

			if (!surnameCell.Value.IsText || string.IsNullOrEmpty(surnameCell.Value.GetText()))
				return null;

			return new()
			{
				Surname = worksheet.Cell(rowIndex, StructureExcel.Student.ColumnNumber.Surname),
				Name = worksheet.Cell(rowIndex, StructureExcel.Student.ColumnNumber.Name),
				MiddleName = worksheet.Cell(rowIndex, StructureExcel.Student.ColumnNumber.MiddleName),
				Snils = worksheet.Cell(rowIndex, StructureExcel.Student.ColumnNumber.Snils),
				INN = worksheet.Cell(rowIndex, StructureExcel.Student.ColumnNumber.Inn),
				Email = worksheet.Cell(rowIndex, StructureExcel.Student.ColumnNumber.Email),
				ExternalId = worksheet.Cell(rowIndex, StructureExcel.Student.ColumnNumber.ExternalId),
				StudyYear = worksheet.Cell(rowIndex, StructureExcel.Student.ColumnNumber.StudyYear),
				Id1 = worksheet.Cell(rowIndex, StructureExcel.Student.ColumnNumber.Id1),
				Title1 = worksheet.Cell(rowIndex, StructureExcel.Student.ColumnNumber.Title1),
				StudyPlan = worksheet.Cell(rowIndex, StructureExcel.Student.ColumnNumber.StudyPlan),
				ID2 = worksheet.Cell(rowIndex, StructureExcel.Student.ColumnNumber.Id2),
			};
		}

		public static Student ToStudent(this TupleCells tuple)
		{
			return new Student
			{
				Surname = tuple.Surname.Value.GetText(),
				Name = tuple.Name.Value.GetText(),
				MiddleName = tuple.MiddleName.Value.GetText(),
				Snils = tuple.Snils.Value.GetText(),
				Inn = tuple.INN.Value.IsNumber? tuple.INN.Value.GetNumber().ToString() : tuple.INN.Value.GetText(),
				Email = tuple.Email.Value.GetText(),
				ExternalId = tuple.ExternalId.Value.GetText(),
				StudyYear = (int)tuple.StudyYear.Value.GetNumber()
			};
		}

		public sealed class TupleCells
		{
			public IXLCell Surname { get; init; }
			public IXLCell Name { get; init; }
			public IXLCell MiddleName { get; init; }
			public IXLCell Snils { get; init; }
			public IXLCell INN { get; init; }
			public IXLCell Email { get; init; }
			public IXLCell ExternalId { get; init; }
			public IXLCell StudyYear { get; init; }
			public IXLCell Id1 { get; init; }
			public IXLCell Title1 { get; init; }
			public IXLCell StudyPlan { get; init; }
			public IXLCell ID2 { get; init; }
		}

		public static void CheckCellType(this TupleCells tuple, int rowIndex)
		{
			if (tuple.Surname is not null && !tuple.Surname.Value.IsText)
				throw new Exception(
					$"Excel of EducationalPrograms, row {rowIndex} Surname is {tuple.Surname.Value.Type}, should be text");

			if (tuple.Name is not null && !tuple.Name.Value.IsText)
				throw new Exception(
					$"Excel of EducationalPrograms, row {rowIndex} Name is {tuple.Name.Value.Type}, should be text");

			if (tuple.MiddleName is not null && !tuple.MiddleName.Value.IsText)
				throw new Exception(
					$"Excel of EducationalPrograms, row {rowIndex} MiddleName is {tuple.MiddleName.Value.Type}, should be text");

			if (tuple.Snils is not null && !tuple.Snils.Value.IsText)
				throw new Exception(
					$"Excel of EducationalPrograms, row {rowIndex} Snils is {tuple.Snils.Value.Type}, should be text");

			if (tuple.INN is not null && !tuple.INN.Value.IsText && !tuple.INN.Value.IsNumber)
				throw new Exception(
					$"Excel of EducationalPrograms, row {rowIndex} INN is {tuple.INN.Value.Type}, should be text or number");

			if (tuple.Email is not null && !tuple.Email.Value.IsText)
				throw new Exception(
					$"Excel of EducationalPrograms, row {rowIndex} Email is {tuple.Email.Value.Type}, should be text");

			if (tuple.ExternalId is not null && !tuple.ExternalId.Value.IsText)
				throw new Exception(
					$"Excel of EducationalPrograms, row {rowIndex} ExternalId is {tuple.ExternalId.Value.Type}, should be text");

			if (tuple.StudyYear is not null && !tuple.StudyYear.Value.IsNumber)
				throw new Exception(
					$"Excel of EducationalPrograms, row {rowIndex} Study Year is {tuple.StudyYear.Value.Type}, should be number");

			if (tuple.StudyYear is not null && !tuple.StudyYear.Value.IsNumber)
				throw new Exception(
					$"Excel of EducationalPrograms, row {rowIndex} Study Year is {tuple.StudyYear.Value.Type}, should be number");
		}

		public static void Check(this Student student, int rowIndex)
		{
			if (string.IsNullOrEmpty(student.Surname))
				throw new Exception($"Excel file, column {rowIndex} Surname is empty");

			if (string.IsNullOrEmpty(student.Name))
				throw new Exception($"Excel file, column {rowIndex} Name is empty");

			if (string.IsNullOrEmpty(student.MiddleName))
				throw new Exception($"Excel file, column {rowIndex} Middle Name is empty");

			if (string.IsNullOrEmpty(student.Snils))
				throw new Exception($"Excel file, column {rowIndex} Snils is empty");

			if (string.IsNullOrEmpty(student.Inn))
				throw new Exception($"Excel file, column {rowIndex} INN is empty");

			if (string.IsNullOrEmpty(student.Email))
				throw new Exception($"Excel file, column {rowIndex} Email is empty");

			if (student.StudyYear <= 0)
				throw new Exception($"Excel file, column {rowIndex} Study Year <= 0");

		}
	}
}
