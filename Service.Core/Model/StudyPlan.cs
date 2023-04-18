using System.Text.Json.Serialization;

namespace Service.Core.Model;

public record class StudyPlan
{
	[JsonPropertyName("id")]
	public string? Id { get; set; }
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
	[JsonPropertyName("education_form")]
	public string? EducationForm { get; set; }
	[JsonPropertyName("educational_program")]
	public string? EducationalProgram { get; set; }
	[JsonPropertyName("external_id")]
	public string? ExternalId { get; set; }
}