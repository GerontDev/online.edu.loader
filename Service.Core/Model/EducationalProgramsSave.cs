using System.Text.Json.Serialization;

namespace Service.Core.Model;

public record EducationalProgramsSave : BaseSave
{
	[JsonPropertyName("educational_programs")]
	public IReadOnlyCollection<EducationalProgram> Disciplines { get; set; } = null!;
}