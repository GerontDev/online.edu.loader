using System.Net.Http.Json;
using System.Text.Encodings.Web;
using System.Text.Json;
using ClosedXML.Excel;
using Service.Core;
using Service.Core.Model;

namespace LoadOfDisciplinesAndStudyPlans;

public class Program
{
	public static async Task Main(string[] args)
	{
		if (!args.Any() || args.Length != 6)
		{
			Console.WriteLine($"Usage: {nameof(LoadOfDisciplinesAndStudyPlans)}  [X-CN-UUID] [OrganizationId] \"[Path to file.xlsx]\" [url to api of online.edu.ru]");
			return;
		}

		string X_CN_UUID = args[0];
		string organizationId = args[1];
		var excelFile = args[2];

		string directoryStudyPlansExcelFile = args[3];
		string directoryDisciplinesExcelFile = args[4];

		string urlOfonline_edu_ru = args[5];
		
		var excelPath = File.Exists(excelFile) ? excelFile : @"..\..\..\..\" + excelFile;
		var directoryStudyPlansExcelPath = File.Exists(directoryStudyPlansExcelFile) ? directoryStudyPlansExcelFile : @"..\..\..\..\" + directoryStudyPlansExcelFile;
		var directoryDisciplinesExcelPath = File.Exists(directoryDisciplinesExcelFile) ? directoryDisciplinesExcelFile : @"..\..\..\..\" + directoryDisciplinesExcelFile;

		Console.WriteLine("Check");

		if (!File.Exists(excelPath))
			throw new FileNotFoundException("File not found", excelPath);

		if (!File.Exists(directoryStudyPlansExcelPath))
			throw new FileNotFoundException("Directory file not found", directoryStudyPlansExcelFile);

		if (!File.Exists(directoryDisciplinesExcelPath))
			throw new FileNotFoundException("Directory file not found", directoryDisciplinesExcelFile);

		if (string.IsNullOrEmpty(X_CN_UUID))
			throw new ArgumentNullException(nameof(X_CN_UUID));

		if (string.IsNullOrEmpty(organizationId))
			throw new ArgumentNullException(nameof(organizationId));

		if (string.IsNullOrEmpty(urlOfonline_edu_ru))
			throw new ArgumentNullException("url to api of online.edu.ru");

		using var workbook = new XLWorkbook(excelPath);
		IXLWorksheet? worksheet = workbook.Worksheets.First();

		HttpClient onlineEduClient = new HttpClient();
		onlineEduClient.DefaultRequestHeaders.Add("X-CN-UUID", X_CN_UUID);
		onlineEduClient.BaseAddress = new Uri(urlOfonline_edu_ru);

		StructureExcel.CheckHeaderExcel(worksheet, StructureExcel.LinkDisciplinesAndStudyPlans.HeaderColumns);

		Console.WriteLine("Run process");

		await ProcessLoading(worksheet,
			onlineEduClient,
			organizationId, 
			StudyPlansExcel.Load(directoryStudyPlansExcelPath),
			DisciplinesExcel.Load(directoryDisciplinesExcelPath));

		workbook.Save();
		Console.WriteLine("Complied");
	}

	private static async Task ProcessLoading(IXLWorksheet worksheet, 
		HttpClient onlineEduClient, 
		string organizationId, 
		IReadOnlyList<StudyPlan> studyPlans,
		IReadOnlyList<Discipline> disciplines)
	{
		for (int rowIndex = 3; rowIndex < ushort.MaxValue; rowIndex++)
		{
			var titleCell = worksheet.Cell(rowIndex, StructureExcel.LinkDisciplinesAndStudyPlans.Columns.Title1ColumnNumber);

			if (!titleCell.Value.IsText || string.IsNullOrEmpty(titleCell.Value.GetText()))
				break;

			string title1 = titleCell.Value.GetText();

			var title1Cell = worksheet.Cell(rowIndex, StructureExcel.LinkDisciplinesAndStudyPlans.Columns.Title1ColumnNumber);
			var title2Cell = worksheet.Cell(rowIndex, StructureExcel.LinkDisciplinesAndStudyPlans.Columns.Title2ColumnNumber);
			var studyPlanCell = worksheet.Cell(rowIndex, StructureExcel.LinkDisciplinesAndStudyPlans.Columns.StudyPlanColumnNumber);
			var disciplineCell = worksheet.Cell(rowIndex, StructureExcel.LinkDisciplinesAndStudyPlans.Columns.DisciplineColumnNumber);
			var semesterCell = worksheet.Cell(rowIndex, StructureExcel.LinkDisciplinesAndStudyPlans.Columns.SemesterColumnNumber);
			var idCell = worksheet.Cell(rowIndex, StructureExcel.LinkDisciplinesAndStudyPlans.Columns.IDColumnNumber);

			DisciplinesAndStudyPlansExcel.CheckCellType(rowIndex, title1Cell, title2Cell, semesterCell);

			string title2 = title2Cell.Value.GetText();
			var discipline = disciplines.Where(_ => string.Equals(_.Title, title2, StringComparison.CurrentCultureIgnoreCase)).SingleOrDefault();
			var studyPlan = studyPlans.Where(_ => string.Equals(_.Title, title1, StringComparison.CurrentCultureIgnoreCase)).SingleOrDefault();

			if (discipline is null)
				throw new Exception($"Don't exist discipline of title2=\"{title2}\". Row {rowIndex}");

			if (string.IsNullOrEmpty(discipline.ExternalId))
				throw new Exception($"Don't exist discipline of title2=\"{title2}\". Row {rowIndex}");

			if (studyPlan is null)
				throw new Exception($"Don't exist discipline of title1=\"{title1}\". Row {rowIndex}");

			if (string.IsNullOrEmpty(studyPlan.ExternalId))
				throw new Exception($"Don't exist discipline of title1=\"{title1}\". Row {rowIndex}");



			var savingLinkDisciplinesAndStudyPlans = new LinkDisciplinesAndStudyPlans()
			{
				DisciplineExternalId = discipline.ExternalId,
				StudyPlanExternalId = studyPlan.ExternalId,
				Semester = (int)semesterCell.Value.GetNumber()
			};

			savingLinkDisciplinesAndStudyPlans.Check(rowIndex);

			var result = await PostEducationalProgramsAsync(client: onlineEduClient, organizationId, savingLinkDisciplinesAndStudyPlans);

			idCell.Value = result.Id;
			studyPlanCell.Value = savingLinkDisciplinesAndStudyPlans.StudyPlanExternalId;
			disciplineCell.Value = savingLinkDisciplinesAndStudyPlans.DisciplineExternalId;
		}
	}

	public static async Task<ResultStatusSave> PostEducationalProgramsAsync(HttpClient client, string organizationId, LinkDisciplinesAndStudyPlans linkDisciplinesAndStudyPlans)
	{
		var educationalPrograms = new LinkDisciplinesAndStudyPlansSave()
		{
			OrganizationId = organizationId,
			StudyPlanDisciplines = new[] { linkDisciplinesAndStudyPlans }
		};
		string jsonString = JsonSerializer.Serialize(educationalPrograms, new JsonSerializerOptions() { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping, });

		var response = await client.PostAsJsonAsync("study_plans_disciplines", educationalPrograms, new JsonSerializerOptions() { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping }, new CancellationToken());

		if (!response.IsSuccessStatusCode)
		{
			Console.WriteLine($"{response.StatusCode} ({(int)response.StatusCode})");
			throw new Exception(await response.Content.ReadAsStringAsync());
		}

		await using var stream = await response.Content.ReadAsStreamAsync();
		var resultObject = JsonSerializer.Deserialize<SaveResults<ResultStatusSave>>(stream);
		ResultStatusSave? infoSave = resultObject?.Results?.FirstOrDefault();

		if (infoSave is null)
		{
			throw new ArgumentNullException(await response.Content.ReadAsStringAsync());
		}

		Console.WriteLine($"\"{linkDisciplinesAndStudyPlans.DisciplineExternalId}\", \"{linkDisciplinesAndStudyPlans.StudyPlanExternalId}\", \"{linkDisciplinesAndStudyPlans.Semester}\"status: {infoSave.UploadStatusType} ({infoSave.AdditionalInfo})");
		return infoSave;
	}

}