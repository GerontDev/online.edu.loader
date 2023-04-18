using ClosedXML.Excel;
using Service.Core.Model;

namespace Service.Core
{
	public static class DisciplinesAndStudyPlansExcel
	{
		public static void CheckCellType(int rowIndex, IXLCell? title1Cell, IXLCell? title2Cell, IXLCell? semesterCell)
		{
			if (!title1Cell.Value.IsText)
				throw new Exception(
					$"Excel of DisciplinesAndStudyPlans, row {rowIndex} Title1 is {title1Cell.Value.Type}, should be text");

			if (!title2Cell.Value.IsText)
				throw new Exception(
					$"Excel of DisciplinesAndStudyPlans, row {rowIndex} Title2 is {title2Cell.Value.Type}, should be text");

			if (!semesterCell.Value.IsNumber)
				throw new Exception(
					$"Excel of DisciplinesAndStudyPlans, row {rowIndex} semester is {semesterCell.Value.Type}, should be number");
		}

		public static void Check(this LinkDisciplinesAndStudyPlans linkDisciplinesAndStudyPlans, int rowIndex)
		{
			if (string.IsNullOrEmpty(linkDisciplinesAndStudyPlans.DisciplineExternalId))
				throw new Exception($"Excel file, row {rowIndex} DisciplineExternalId is empty");

			if (string.IsNullOrEmpty(linkDisciplinesAndStudyPlans.StudyPlanExternalId))
				throw new Exception($"Excel file, row {rowIndex} StudyPlanExternalId is empty");

			if (linkDisciplinesAndStudyPlans.Semester <= 0)
				throw new Exception($"Excel file, row {rowIndex} Semester <= 0 ");
		}
	}
}
