using System.Text.Json.Serialization;

namespace Service.Core.Model;

public record class EducationalProgram
{
	[JsonPropertyName("id")]
	public string? Id { get; set; }
	[JsonPropertyName("external_id")]
	public string? ExternalId { get; set; }
	[JsonPropertyName("title")]
	public string? Title { get; set; }
	[JsonPropertyName("direction")]
	public string? Direction { get; set; }
	[JsonPropertyName("code_direction")]
	public string? CodeDirection { get; set; }
	[JsonPropertyName("start_year")]
	public int StartYear { get; set; }
	[JsonPropertyName("end_year")]
	public int EndYear { get; set; }
}
