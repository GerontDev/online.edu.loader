using System.Text.Json.Serialization;

namespace Service.Core.Model
{
	public record LinkDisciplinesAndStudyPlansSave : BaseSave
	{
		[JsonPropertyName("study_plan_disciplines")]
		public IReadOnlyCollection<LinkDisciplinesAndStudyPlans> StudyPlanDisciplines { get; set; } = null!;
	}
}
