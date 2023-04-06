using ClosedXML.Excel;
using Service.Core.Model;
using System.Net.Http.Json;
using System.Text.Encodings.Web;
using System.Text.Json;

namespace LoadOfStudyPlans;

public class Program
{
	private const string StudyPansPrefix = "STPL";

	public static async Task Main(string[] args)
	{
		if (!args.Any() || args.Length != 4)
		{
			Console.WriteLine("Usage: LoadOfEducationalPrograms [X-CN-UUID] [OrganizationId] \"[Path to file.xlsx]\" [url to api of online.edu.ru]");
			return;
		}

		string X_CN_UUID = args[0];
		string organizationId = args[1];
		var excelFile = args[2];
		string urlOfonline_edu_ru = args[3];

		var excelPath = File.Exists(excelFile) ? excelFile : @"..\..\..\..\" + excelFile;

		Console.WriteLine("Check");

		if (!File.Exists(excelPath))
			throw new FileNotFoundException("File not found", excelPath);

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

		var checkColumns = new[] { "external_id", "title", "direction", "code_direction", "start_year", "end_year", "education_form", "educational_program", "ID" };

		if (IsInvalidExcelFile(worksheet, checkColumns))
			throw new Exception($"File is invalid. Columns is need [{string.Join(", ", checkColumns)}] ");


		Console.WriteLine("Run process");

		await ProcessLoading(worksheet, onlineEduClient, organizationId);

		workbook.Save();
		Console.WriteLine("Complied");
	}

	private static async Task ProcessLoading(IXLWorksheet worksheet, HttpClient onlineEduClient, string organizationId)
	{
		const int externalIdColumnNumber = 1;
		const int tileColumnNumber = 2;
		const int directionColumnNumber = 3;
		const int codeDirectionColumnNumber = 4;
		const int startYearColumnNumber = 5;
		const int endYearColumnNumber = 6;
		const int educationFormColumnNumber = 7;
		const int educationalProgramColumnNumber = 8;
		const int idColumnNumber = 9;

		for (int r = 3; r < ushort.MaxValue; r++)
		{
			var titleCell = worksheet.Cell(r, tileColumnNumber);

			if (!titleCell.Value.IsText || string.IsNullOrEmpty(titleCell.Value.GetText()))
				break;

			var externalIdCell = worksheet.Cell(r, externalIdColumnNumber);
			var directionCell = worksheet.Cell(r, directionColumnNumber);
			var codeDirectionCell = worksheet.Cell(r, codeDirectionColumnNumber);
			var startYearCell = worksheet.Cell(r, startYearColumnNumber);
			var endYearCall = worksheet.Cell(r, endYearColumnNumber);
			var educationFormCall = worksheet.Cell(r, educationFormColumnNumber);
			var educationalProgramCall = worksheet.Cell(r, educationalProgramColumnNumber);


			string externalId = StudyPansPrefix + Guid.NewGuid().ToString();

			if (!externalIdCell.Value.IsText || string.IsNullOrEmpty(externalIdCell.Value.GetText()))
			{
				worksheet.Cell(r, 1).Value = externalId;
			}
			else
			{
				externalId = externalIdCell.Value.GetText();
				worksheet.Cell(r, 1).Value = externalId;
			}

			if (codeDirectionCell.Value.IsDateTime)
				throw new Exception($"Excel file, row {r} Code Direction is DataTime, should be text");

			var savingStudyPlan = new StudyPlan()
			{
				ExternalId = externalId,
				Title = titleCell.Value.GetText(),
				Direction = directionCell.Value.GetText(),
				CodeDirection = codeDirectionCell.Value.GetText(),
				EndYear = (int)startYearCell.Value.GetNumber(),
				StartYear = (int)endYearCall.Value.GetNumber(),
				EducationForm = educationFormCall.Value.GetText(),
				EducationalProgram = educationalProgramCall.Value.GetText()
			};

			Check(savingStudyPlan, r);

			var result = await PostStudyPlansAsync(client: onlineEduClient, organizationId, savingStudyPlan);

			worksheet.Cell(r, idColumnNumber).Value = result.Id;
		}
	}

	private static void Check(StudyPlan studyPlan, int rowIndex)
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

	private static bool IsInvalidExcelFile(IXLWorksheet worksheet, string[] columns)
	{
		for (int c = 1; c <= columns.Length; c++)
		{
			if (worksheet.Cell(c, 1).Value.GetText() != columns[c - 1])
				return false;
		}

		return true;
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