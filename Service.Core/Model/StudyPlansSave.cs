using System.Text.Json.Serialization;

namespace Service.Core.Model;
public record class StudyPlansSave : BaseSave
{
	[JsonPropertyName("study_plans")]
	public IReadOnlyCollection<StudyPlan> StudyPlans { get; set; } = null!;
}