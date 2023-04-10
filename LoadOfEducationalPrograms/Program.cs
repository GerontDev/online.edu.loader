using ClosedXML.Excel;
using Service.Core;
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

		StructureExcel.CheckHeaderExcel(worksheet, StructureExcel.EducationalPrograms.HeaderColumns);

		Console.WriteLine("Run process");

		await ProcessLoading(worksheet, onlineEduClient, organizationId);

		workbook.Save();
		Console.WriteLine("Complied");
	}

	private static async Task ProcessLoading(IXLWorksheet worksheet, HttpClient onlineEduClient, string organizationId)
	{
		for (int rowIndex = 3; rowIndex < ushort.MaxValue; rowIndex++)
		{
			var titleCell = worksheet.Cell(rowIndex, StructureExcel.EducationalPrograms.Columns.TileColumnNumber);

			if (!titleCell.Value.IsText || string.IsNullOrEmpty(titleCell.Value.GetText()))
				break;

			string title = titleCell.Value.GetText();

			var externalIdCell = worksheet.Cell(rowIndex, StructureExcel.EducationalPrograms.Columns.ExternalIdColumnNumber);
			var directionCell = worksheet.Cell(rowIndex, StructureExcel.EducationalPrograms.Columns.DirectionColumnNumber);
			var codeDirectionCell = worksheet.Cell(rowIndex, StructureExcel.EducationalPrograms.Columns.CodeDirectionColumnNumber);
			var startYearCell = worksheet.Cell(rowIndex, StructureExcel.EducationalPrograms.Columns.StartYearColumnNumber);
			var endYearCall = worksheet.Cell(rowIndex, StructureExcel.EducationalPrograms.Columns.EndYearColumnNumber);

			

			string externalId = EducationalProgramPrefix + Guid.NewGuid().ToString();

			if (!externalIdCell.Value.IsText || string.IsNullOrEmpty(externalIdCell.Value.GetText()))
			{
				worksheet.Cell(rowIndex, 1).Value = externalId;
			}
			else
			{
				externalId = externalIdCell.Value.GetText();
				worksheet.Cell(rowIndex, 1).Value = externalId;
			}

			EducationalProgramsExcel.CheckCellType(rowIndex, null, directionCell, codeDirectionCell, startYearCell, endYearCall, null);

			var savingEducationalProgram = new EducationalProgram()
			{
				ExternalId = externalId,
				Title = title,
				Direction = directionCell.Value.GetText(),
				CodeDirection = codeDirectionCell.Value.GetText(),
				EndYear = (int)startYearCell.Value.GetNumber(),
				StartYear = (int)endYearCall.Value.GetNumber()
			};

			savingEducationalProgram.Check(rowIndex);

			var result = await PostEducationalProgramsAsync(client: onlineEduClient, organizationId, savingEducationalProgram);

			worksheet.Cell(rowIndex, StructureExcel.EducationalPrograms.Columns.IdColumnNumber).Value = result.Id;
		}
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