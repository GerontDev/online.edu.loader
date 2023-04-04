using System.Text.Json.Serialization;

namespace LoadOfDisciplines.Models
{
	public class ResultsPost<T>
	{
		[JsonPropertyName("results")]
		public IReadOnlyCollection<T>? Results { get; init; }
	}
}
