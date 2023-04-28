using System.Text.Json.Serialization;

namespace Service.Core.Model;

public record class Student
{
	[JsonPropertyName("id")]
	public string? Id { get; init; }
	[JsonPropertyName("surname")]
	public string Surname { get; init; } = string.Empty;
	[JsonPropertyName("name")]
	public string Name { get; init; } = string.Empty;
	[JsonPropertyName("middle_name")]
	public string MiddleName { get; init; } = string.Empty;
	[JsonPropertyName("snils")]
	public string Snils { get; init; } = string.Empty;
	[JsonPropertyName("inn")]
	public string Inn { get; init; } = string.Empty;
	[JsonPropertyName("email")]
	public string Email { get; init; } = string.Empty;
	[JsonPropertyName("external_id")]
	public string ExternalId { get; set; } = string.Empty;
	[JsonPropertyName("study_year")]
	public int StudyYear { get; init; }
}
