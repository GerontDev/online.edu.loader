using System.Text.Json.Serialization;

namespace Service.Core.Model
{
	public class ResultsPost<T>
	{
		[JsonPropertyName("results")]
		public IReadOnlyCollection<T>? Results { get; init; }
	}
}
