using System.Text.Json.Serialization;

namespace Service.Core.Model;

public class SaveResults<T>
{
	[JsonPropertyName("results")]
	public IReadOnlyCollection<T>? Results { get; init; }
}
