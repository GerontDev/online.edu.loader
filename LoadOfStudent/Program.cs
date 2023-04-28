using System.Net.Http.Json;
using System.Text.Encodings.Web;
using System.Text.Json;
using ClosedXML.Excel;
using Service.Core;
using Service.Core.Model;

namespace LoadOfStudent;

public class Program
{
	private const string StudentPrefix = "ST";

	public static async Task Main(string[] args)
	{
		if (!args.Any() || args.Length != 4)
		{
			Console.WriteLine($"Usage: {nameof(LoadOfStudent)} [X-CN-UUID] [OrganizationId] \"[Path to file.xlsx]\" [url to api of online.edu.ru]");
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

		StructureExcel.CheckHeaderExcel(worksheet, StructureExcel.Student.HeaderColumns);

		Console.WriteLine("Run process");

		await ProcessLoading(worksheet, onlineEduClient, organizationId);

		workbook.Save();
		Console.WriteLine("Complied");
	}

	private static async Task ProcessLoading(IXLWorksheet worksheet, HttpClient onlineEduClient, string organizationId)
	{
		for (int rowIndex = 3; rowIndex < ushort.MaxValue; rowIndex++)
		{
			StudentExcel.TupleCells? tuple = worksheet.GetTupleCells(rowIndex);
			if (tuple is null)
				break;

			string externalId = StudentPrefix + Guid.NewGuid();
			if (!tuple.ExternalId.Value.IsText || string.IsNullOrEmpty(tuple.ExternalId.Value.GetText()))
			{
				tuple.ExternalId.Value = externalId;
			}
			else
			{
				externalId = tuple.ExternalId.Value.GetText();
			}
			tuple.CheckCellType(rowIndex);
			var savingStudent = tuple.ToStudent();
			savingStudent.ExternalId = externalId;
			savingStudent.Check(rowIndex);

			var result = await PostStudentAsync(client: onlineEduClient, organizationId, savingStudent);
			tuple.Id1.Value = result.Id;
		}
	}

	public static async Task<ResultStatusSave> PostStudentAsync(HttpClient client, string organizationId, Student student)
	{
		var educationalPrograms = new StudentSave()
		{
			OrganizationId = organizationId,
			Students = new[] { student }
		};
		string jsonString = JsonSerializer.Serialize(educationalPrograms, new JsonSerializerOptions() { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping, });

		var response = await client.PostAsJsonAsync("students", educationalPrograms, new JsonSerializerOptions() { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping }, new CancellationToken());

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

		Console.WriteLine($"\"{student.Surname}\" \"{student.Name}\" \"{student.MiddleName}\" status: {infoSave.UploadStatusType} ({infoSave.AdditionalInfo})");
		return infoSave;
	}

}