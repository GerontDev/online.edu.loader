using System.Text.Json.Serialization;

namespace Service.Core.Model;

public record class Discipline
{
	[JsonPropertyName("id")]
	public string? Id { get; set; }
	[JsonPropertyName("title")]
	public string Title { get; set; } = null!;
	[JsonPropertyName("external_id")]
	public string ExternalId { get; set; } = null!;
}