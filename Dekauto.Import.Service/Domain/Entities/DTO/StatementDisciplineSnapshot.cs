namespace Dekauto.Import.Service.Domain.Entities
{
    /// <summary>Снимок дисциплины из ведомости до сопоставления с учебным планом.</summary>
    public sealed class StatementDisciplineSnapshot
    {
        public string StatementName { get; init; } = string.Empty;
        public string Score { get; init; } = string.Empty;
        public short? Semester { get; init; }
        public short? Year { get; init; }
        public bool IsCourseWork { get; init; }
        public bool IsPractice { get; init; }
        public string? SourceFileName { get; init; }
        public string? SourceSheetName { get; init; }

        public string MergeKey =>
            IsCourseWork
                ? $"cw:{Semester}"
                : $"{Semester}:{StatementName.Trim().ToLowerInvariant()}";
    }
}
