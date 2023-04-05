using System.Text.Json.Serialization;

namespace Service.Core.Model;

public record class ResultStatusSave
{
	[JsonPropertyName("id")]
	public string? Id { get; init; }
	[JsonPropertyName("external_id")]
	public string? ExternalId { get; init; }
	[JsonPropertyName("uploadStatusType")]
	public string? UploadStatusType { get; init; }
	[JsonPropertyName("additional_info")]
	public string? AdditionalInfo { get; init; }
}
