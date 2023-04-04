using System.Text.Json.Serialization;

namespace LoadOfDisciplines.Models;

public record class Discipline
{
	[JsonPropertyName("title")]
	public string Title { get; set; } = null!;

	[JsonPropertyName("external_id")]
	public string ExternalId { get; set; } = null!;
}