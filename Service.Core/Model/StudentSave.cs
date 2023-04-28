using System.Text.Json.Serialization;

namespace Service.Core.Model;

public record StudentSave : BaseSave
{
	[JsonPropertyName("students")]
	public IReadOnlyCollection<Student> Students{ get; set; } = null!;
}