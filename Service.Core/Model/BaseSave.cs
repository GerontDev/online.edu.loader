using System.Text.Json.Serialization;

namespace Service.Core.Model;

public record class BaseSave
{
	[JsonPropertyName("organization_id")]
	public string OrganizationId { get; set; } = null!;
}
