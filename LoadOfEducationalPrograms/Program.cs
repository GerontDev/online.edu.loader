using ClosedXML.Excel;
using Service.Core.Model;
using System.Net.Http.Json;
using System.Text.Encodings.Web;
using System.Text.Json;

namespace LoadOfEducationalPrograms;

public class Program
{
	private const string EducationalProgramPrefix = "EDUPRO";

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

		var checkColumns = new[] { "external_id", "title", "direction", "code_direction", "start_year", "end_year", "ID" };

		if (IsInvalideExcelFile(worksheet, checkColumns))
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
		const int idColumnNumber = 7;

		for (int r = 3; r < ushort.MaxValue; r++)
		{
			var titleCell = worksheet.Cell(r, tileColumnNumber);

			if (!titleCell.Value.IsText || string.IsNullOrEmpty(titleCell.Value.GetText()))
				break;

			string title = titleCell.Value.GetText();

			var externalIdCell = worksheet.Cell(r, externalIdColumnNumber);
			var directionCell = worksheet.Cell(r, directionColumnNumber);
			var codeDirectionCell = worksheet.Cell(r, codeDirectionColumnNumber);
			var startYearCell = worksheet.Cell(r, startYearColumnNumber);
			var endYearCall = worksheet.Cell(r, endYearColumnNumber);


			string externalId = EducationalProgramPrefix +  Guid.NewGuid().ToString();

			if (!externalIdCell.Value.IsText || string.IsNullOrEmpty(externalIdCell.Value.GetText()))
			{
				worksheet.Cell(r, 1).Value = externalId;
			}
			else
			{
				externalId = externalIdCell.Value.GetText();
				worksheet.Cell(r, 1).Value = externalId;
			}

			var direction = directionCell.Value.GetText();
			if (string.IsNullOrEmpty(direction))
				throw new Exception($"Excel file, column {r} direction is empty");

			var codeDirection = codeDirectionCell.Value.GetText();
			if (string.IsNullOrEmpty(codeDirection))
				throw new Exception($"Excel file, column {r} Code Direction is empty");

			var startYear = startYearCell.Value.GetNumber();
			if (startYear < 1900)
				throw new Exception($"Excel file, column {r} Code Direction is invalid");

			var endYear = endYearCall.Value.GetNumber();
			if (endYear < 1900)
				throw new Exception($"Excel file, column {r} Code Direction is invalid");

			var savingEducationalProgram = new EducationalProgram()
			{
				ExternalId = externalId,
				Title = title,
				Direction = direction,
				CodeDirection = codeDirection,
				EndYear = (int)endYear,
				StartYear = (int)startYear
			};
			var result = await PostEducationalProgramsAsync(client: onlineEduClient, organizationId, savingEducationalProgram);

			worksheet.Cell(r, idColumnNumber).Value = result.Id;
		}
	}

	private static bool IsInvalideExcelFile(IXLWorksheet worksheet, string[] columns)
	{
		for (int c = 1; c <= columns.Length; c++)
		{
			if (worksheet.Cell(c, 1).Value.GetText() != columns[c - 1])
				return false;
		}

		return true;
	}

	public static async Task<ResultStatusSave> PostEducationalProgramsAsync(HttpClient client, string organizationId, EducationalProgram educationalProgram)
	{
		var educationalPrograms = new EducationalProgramsSave()
		{
			OrganizationId = organizationId,
			Disciplines = new[] { educationalProgram }
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

		Console.WriteLine($"\"{educationalProgram.Title}\" status: {infoSave.UploadStatusType} ({infoSave.AdditionalInfo})");
		return infoSave;
	}

}