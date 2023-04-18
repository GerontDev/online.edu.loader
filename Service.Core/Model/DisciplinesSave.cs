using System.Text.Json.Serialization;

namespace Service.Core.Model;

public record class DisciplinesSave
{
	[JsonPropertyName("organization_id")]
	public string OrganizationId { get; set; } = null!;

	[JsonPropertyName("disciplines")]
	public IReadOnlyCollection<Discipline> Disciplines { get; set; } = null!;
}
