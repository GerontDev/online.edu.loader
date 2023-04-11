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
		if (!args.Any() || args.Length != 4)
		{
			Console.WriteLine($"Usage: {nameof(LoadOfDisciplinesAndStudyPlans)}  [X-CN-UUID] [OrganizationId] \"[Path to file.xlsx]\" [url to api of online.edu.ru]");
			return;
		}

		string X_CN_UUID = args[0];
		string organizationId = args[1];
		var excelFile = args[2];
		string directoryEducationalProgramsExcelFile = args[3];
		string directoryDisciplineExcelFile = args[4];
		string urlOfonline_edu_ru = args[5];

		var excelPath = File.Exists(excelFile) ? excelFile : @"..\..\..\..\" + excelFile;

		Console.WriteLine("Check");

		if (!File.Exists(excelPath))
			throw new FileNotFoundException("File not found", excelPath);

		if (!File.Exists(directoryEducationalProgramsExcelFile))
			throw new FileNotFoundException("Directory file not found", directoryEducationalProgramsExcelFile);

		if (!File.Exists(directoryEducationalProgramsExcelFile))
			throw new FileNotFoundException("Directory file not found", directoryEducationalProgramsExcelFile);

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

		await ProcessLoading(worksheet, onlineEduClient, organizationId, EducationalProgramsExcel.Load(directoryEducationalProgramsExcelFile));

		workbook.Save();
		Console.WriteLine("Complied");
	}

	private static async Task ProcessLoading(IXLWorksheet worksheet, HttpClient onlineEduClient, string organizationId, IReadOnlyList<EducationalProgram> educationalPrograms)
	{
		for (int rowIndex = 3; rowIndex < ushort.MaxValue; rowIndex++)
		{
			var titleCell = worksheet.Cell(rowIndex, StructureExcel.LinkDisciplinesAndStudyPlans.Columns.Title1ColumnNumber);

			if (!titleCell.Value.IsText || string.IsNullOrEmpty(titleCell.Value.GetText()))
				break;

			string title = titleCell.Value.GetText();

			var title1Cell = worksheet.Cell(rowIndex, StructureExcel.LinkDisciplinesAndStudyPlans.Columns.Title1ColumnNumber);
			var title2Cell = worksheet.Cell(rowIndex, StructureExcel.LinkDisciplinesAndStudyPlans.Columns.Title2ColumnNumber);
			var studyPlanCell = worksheet.Cell(rowIndex, StructureExcel.LinkDisciplinesAndStudyPlans.Columns.StudyPlanColumnNumber);
			var disciplineCell = worksheet.Cell(rowIndex, StructureExcel.LinkDisciplinesAndStudyPlans.Columns.DisciplineColumnNumber);
			var semesterCell = worksheet.Cell(rowIndex, StructureExcel.LinkDisciplinesAndStudyPlans.Columns.SemesterColumnNumber);



			if (!title1Cell.Value.IsText)
				throw new Exception(
					$"Excel of DisciplinesAndStudyPlans, row {rowIndex} Title1 is {title1Cell.Value.Type}, should be text");

			if (!title2Cell.Value.IsText)
				throw new Exception(
					$"Excel of DisciplinesAndStudyPlans, row {rowIndex} Title2 is {title2Cell.Value.Type}, should be text");

			if (!semesterCell.Value.IsText)
				throw new Exception(
					$"Excel of DisciplinesAndStudyPlans, row {rowIndex} semester is {semesterCell.Value.Type}, should be text");


			var savingLinkDisciplinesAndStudyPlans = new LinkDisciplinesAndStudyPlans()
			{
				Title = title,
				Direction = directionCell.Value.GetText(),
				CodeDirection = codeDirectionCell.Value.GetText(),
				EndYear = (int)startYearCell.Value.GetNumber(),
				StartYear = (int)endYearCall.Value.GetNumber()
			};

			savingEducationalProgram.Check(rowIndex);

			var result = await PostEducationalProgramsAsync(client: onlineEduClient, organizationId, savingLinkDisciplinesAndStudyPlans);

			worksheet.Cell(rowIndex, StructureExcel.EducationalPrograms.Columns.IdColumnNumber).Value = result.Id;
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

		var response = await client.PostAsJsonAsync("educational_programs", educationalPrograms, new JsonSerializerOptions() { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping }, new CancellationToken());

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

		Console.WriteLine($"\"{linkDisciplinesAndStudyPlans.Title}\" status: {infoSave.UploadStatusType} ({infoSave.AdditionalInfo})");
		return infoSave;
	}

}