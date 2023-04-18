using ClosedXML.Excel;
using Service.Core.Model;
using System.Net.Http.Json;
using System.Text.Encodings.Web;
using System.Text.Json;
using Service.Core;

namespace LoadOfStudyPlans;

public class Program
{
	private const string StudyPansPrefix = "STPL";

	public static async Task Main(string[] args)
	{
		if (!args.Any() || args.Length != 5)
		{
			Console.WriteLine($"Usage: {nameof(LoadOfStudyPlans)} [X-CN-UUID] [OrganizationId] \"[Path to file.xlsx]\" \"[Path to directory file.xlsx]\" [url to api of online.edu.ru]");
			return;
		}

		string X_CN_UUID = args[0];
		string organizationId = args[1];
		var excelFile = args[2];
		string directoryExcelFile = args[3];
		string urlOfonline_edu_ru = args[4];

		var excelPath = File.Exists(excelFile) ? excelFile : @"..\..\..\..\" + excelFile;
		var directoryExcelPath = File.Exists(directoryExcelFile) ? directoryExcelFile : @"..\..\..\..\" + directoryExcelFile;

		Console.WriteLine("Check");

		if (!File.Exists(excelPath))
			throw new FileNotFoundException("File not found", excelPath);

		if (!File.Exists(directoryExcelPath))
			throw new FileNotFoundException("Directory file not found", directoryExcelFile);

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

		StructureExcel.CheckHeaderExcel(worksheet, StructureExcel.StudyPlans.HeaderColumns);

		Console.WriteLine("Run process");

		await ProcessLoading(worksheet, onlineEduClient, organizationId, EducationalProgramsExcel.Load(directoryExcelPath));

		workbook.Save();
		Console.WriteLine("Complied");
	}

	private static async Task ProcessLoading(IXLWorksheet worksheet, HttpClient onlineEduClient, string organizationId, IReadOnlyList<EducationalProgram> educationalPrograms)
	{

		for (int rowIndex = 3; rowIndex < ushort.MaxValue; rowIndex++)
		{
			var titleCell = worksheet.Cell(rowIndex, StructureExcel.StudyPlans.Columns.TileColumnNumber);

			if (!titleCell.Value.IsText || string.IsNullOrEmpty(titleCell.Value.GetText()))
				break;

			var externalIdCell = worksheet.Cell(rowIndex, StructureExcel.StudyPlans.Columns.ExternalIdColumnNumber);
			var directionCell = worksheet.Cell(rowIndex, StructureExcel.StudyPlans.Columns.DirectionColumnNumber);
			var codeDirectionCell = worksheet.Cell(rowIndex, StructureExcel.StudyPlans.Columns.CodeDirectionColumnNumber);
			var startYearCell = worksheet.Cell(rowIndex, StructureExcel.StudyPlans.Columns.StartYearColumnNumber);
			var endYearCall = worksheet.Cell(rowIndex, StructureExcel.StudyPlans.Columns.EndYearColumnNumber);
			var educationFormCall = worksheet.Cell(rowIndex, StructureExcel.StudyPlans.Columns.EducationFormColumnNumber);
			var educationalProgramCall = worksheet.Cell(rowIndex, StructureExcel.StudyPlans.Columns.EducationalProgramColumnNumber);


			string externalId = StudyPansPrefix + Guid.NewGuid().ToString();

			if (!externalIdCell.Value.IsText || string.IsNullOrEmpty(externalIdCell.Value.GetText()))
			{
				worksheet.Cell(rowIndex, 1).Value = externalId;
			}
			else
			{
				externalId = externalIdCell.Value.GetText();
				worksheet.Cell(rowIndex, 1).Value = externalId;
			}

			StudyPlansExcel.CheckCellType(rowIndex, null, directionCell,codeDirectionCell, startYearCell, endYearCall, educationFormCall, educationalProgramCall, null);

			var savingStudyPlan = new StudyPlan()
			{
				ExternalId = externalId,
				Title = titleCell.Value.GetText(),
				Direction = directionCell.Value.GetText(),
				CodeDirection = codeDirectionCell.Value.GetText(),
				EndYear = (int)startYearCell.Value.GetNumber(),
				StartYear = (int)endYearCall.Value.GetNumber(),
				EducationForm = educationFormCall.Value.GetText(),
				//EducationalProgram = educationalProgramCall.Value.GetText()
			};

			var educationalProgram = educationalPrograms.FirstOrDefault(_ =>
				string.Equals(_.Direction, savingStudyPlan.Direction, StringComparison.CurrentCultureIgnoreCase));

			if (educationalProgram is null)
				throw new Exception($"Don't exist educational program of direction=\"{savingStudyPlan.Direction}\". Row{rowIndex}");

			if (string.IsNullOrEmpty(educationalProgram.ExternalId))
				throw new Exception($"Don't exist educational program of direction=\"{savingStudyPlan.Direction}\". Row{rowIndex}");

			savingStudyPlan.EducationalProgram = educationalProgram.ExternalId;

			savingStudyPlan.Check(rowIndex);

			var result = await PostStudyPlansAsync(client: onlineEduClient, organizationId, savingStudyPlan);

			worksheet.Cell(rowIndex, StructureExcel.StudyPlans.Columns.IdColumnNumber).Value = result.Id;
			worksheet.Cell(rowIndex, StructureExcel.StudyPlans.Columns.EducationalProgramColumnNumber).Value = educationalProgram.ExternalId;
		}
	}


	public static async Task<ResultStatusSave> PostStudyPlansAsync(HttpClient client, string organizationId, StudyPlan studyPlan)
	{
		var educationalPrograms = new StudyPlansSave()
		{
			OrganizationId = organizationId,
			StudyPlans = new[] { studyPlan }
		};
		string jsonString = JsonSerializer.Serialize(educationalPrograms, new JsonSerializerOptions() { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping, });

		var response = await client.PostAsJsonAsync("study_plans", educationalPrograms, new JsonSerializerOptions() { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping }, new CancellationToken());

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

		Console.WriteLine($"\"{studyPlan.Title}\" status: {infoSave.UploadStatusType} ({infoSave.AdditionalInfo})");
		return infoSave;
	}

}