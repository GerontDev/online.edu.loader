using System.Text.Json.Serialization;

namespace Service.Core.Model
{
	public class LinkDisciplinesAndStudyPlans
	{
		[JsonPropertyName("study_plan")]
		public string? StudyPlanExternalId { get; set; }
		[JsonPropertyName("discipline")]
		public string? DisciplineExternalId { get; set; }
		[JsonPropertyName("semester")]
		public int Semester { get; set; }
	}
}
