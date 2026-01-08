namespace Dekauto.Import.Service.Domain.Entities
{
    public class StudentDisciplineResult
    {
        public string? DisciplineName { get; set; }
        public double? Score { get; set; }
        public short? Semester { get; set; }
        public short? Year { get; set; }
        public string? ControlType { get; set; }

        public double? AudHours { get; set; }
        public double? CreditUnits { get; set; }
    }
}
