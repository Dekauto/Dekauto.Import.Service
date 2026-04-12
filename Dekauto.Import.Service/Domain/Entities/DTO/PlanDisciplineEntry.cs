namespace Dekauto.Import.Service.Domain.Entities.DTO
{
    public class PlanDisciplineEntry
    {
        public string DisciplineName { get; set; } = string.Empty;
        public int PlanOrder { get; set; }
        public double? TotalAudHours { get; set; }
        public Dictionary<int, double> CreditUnitsBySemester { get; set; } = new();
    }
}
