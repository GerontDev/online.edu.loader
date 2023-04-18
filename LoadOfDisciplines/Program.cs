using ClosedXML.Excel;
using System.Net.Http.Json;
using System.Text.Encodings.Web;
using System.Text.Json;
using Service.Core;
using Service.Core.Model;

namespace LoadOfDisciplines;

public class Program
{
	private const string DisciplinePrefix = "DISC";

	public static async Task Main(string[] args)
	{
		if (!args.Any() || args.Length != 4)
		{
			Console.WriteLine($"Usage: {nameof(LoadOfDisciplines)} [X-CN-UUID] [OrganizationId] \"[Path to Шаблон дисциплины.xlsx]\" [url to api of online.edu.ru]");
			return;
		}

		string X_CN_UUID = args[0];
		string OrganizationId = args[1];
		var excelFile = args[2];
		string urlOfonline_edu_ru = args[3];

		var excelPath = File.Exists(excelFile) ? excelFile : @"..\..\..\..\" + excelFile;

		Console.WriteLine("Check");

		if (!File.Exists(excelPath))
			throw new FileNotFoundException("File not found", excelPath);

		if (string.IsNullOrEmpty(X_CN_UUID))
			throw new ArgumentNullException(nameof(X_CN_UUID));

		if (string.IsNullOrEmpty(OrganizationId))
			throw new ArgumentNullException(nameof(OrganizationId));

		if (string.IsNullOrEmpty(urlOfonline_edu_ru))
			throw new ArgumentNullException("url to api of online.edu.ru");

		using var workbook = new XLWorkbook(excelPath);
		var worksheet = workbook.Worksheets.First();

		HttpClient onlineEduClient = new HttpClient();
		onlineEduClient.DefaultRequestHeaders.Add("X-CN-UUID", X_CN_UUID);
		onlineEduClient.BaseAddress = new Uri(urlOfonline_edu_ru);
		
		StructureExcel.CheckHeaderExcel(worksheet, StructureExcel.Disciplines.HeaderColumns);


		Console.WriteLine("Run process");
		for (int r = 3; r < ushort.MaxValue; r++)
		{
			var titleCell = worksheet.Cell(r, StructureExcel.Disciplines.Columns.TileColumnNumber);

			if (!titleCell.Value.IsText || string.IsNullOrEmpty(titleCell.Value.GetText()))
				break;

			string title = titleCell.Value.GetText();
			var externalIdCall = worksheet.Cell(r, StructureExcel.Disciplines.Columns.ExternalIdColumnNumber);
			string externalId = DisciplinePrefix + Guid.NewGuid().ToString();

			if (!externalIdCall.Value.IsText || string.IsNullOrEmpty(externalIdCall.Value.GetText()))
			{
				worksheet.Cell(r, StructureExcel.Disciplines.Columns.ExternalIdColumnNumber).Value = externalId;
			}
			else
			{
				externalId = externalIdCall.Value.GetText();
				worksheet.Cell(r, StructureExcel.Disciplines.Columns.ExternalIdColumnNumber).Value = externalId;
			}

			var result = await PostDisciplinesAsync(client: onlineEduClient, OrganizationId,
				new Discipline() { ExternalId = externalId, Title = title });

			worksheet.Cell(r, StructureExcel.Disciplines.Columns.IdColumnNumber).Value = result.Id;
		}

		workbook.Save();
		Console.WriteLine("Complied");
	}

	public static async Task<DisciplineStatusSave> PostDisciplinesAsync(HttpClient client, string organizationId, Discipline discipline)
	{
		var disciplines = new DisciplinesSave()
		{
			OrganizationId = organizationId,
			Disciplines = new[] { discipline }
		};
		string jsonString = JsonSerializer.Serialize(disciplines, new JsonSerializerOptions() { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping, });
		var response = await client.PostAsJsonAsync("disciplines", disciplines, new JsonSerializerOptions() { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping }, new CancellationToken());
		if (!response.IsSuccessStatusCode)
			throw new Exception(await response.Content.ReadAsStringAsync());
		await using var stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false);
		var resultObject = JsonSerializer.Deserialize<ResultsPost<DisciplineStatusSave>>(stream);
		DisciplineStatusSave? infoSave = resultObject?.Results?.FirstOrDefault();
		if (infoSave is null)
			throw new ArgumentNullException(await response.Content.ReadAsStringAsync());
		Console.WriteLine($"\"{discipline.Title}\" status: {infoSave.UploadStatusType} ({infoSave.AdditionalInfo})");
		return infoSave;
	}

}