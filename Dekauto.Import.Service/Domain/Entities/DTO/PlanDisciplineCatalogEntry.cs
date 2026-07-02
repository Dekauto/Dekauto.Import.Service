namespace Dekauto.Import.Service.Domain.Entities
{
    /// <summary>Строка каталога дисциплин из листов «Курс N» учебного плана.</summary>
    public sealed class PlanDisciplineCatalogEntry
    {
        public short Semester { get; init; }
        public string PlanName { get; init; } = string.Empty;
        public string DisplayName { get; init; } = string.Empty;
        public bool IsPractice { get; init; }
        public string? ControlType { get; init; }
        public int? TotalHours { get; init; }
        public int? AudHours { get; init; }
    }
}
