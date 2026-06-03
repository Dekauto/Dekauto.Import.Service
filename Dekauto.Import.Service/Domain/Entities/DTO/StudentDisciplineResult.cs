using Dekauto.Import.Service.Domain.Entities.DTO;

namespace Dekauto.Import.Service.Domain.Entities
{
    public class StudentDisciplineResult
    {
        /// <summary>Из учебного плана; если null — только карточка или ГИА/прочее.</summary>
        public SupplementPlanBucket? PlanBucket { get; set; }
        public string? DisciplineName { get; set; }
        public string? Score { get; set; }
        public short? Semester { get; set; }
        public short? Year { get; set; }
        public string? ControlType { get; set; }

        public double? AudHours { get; set; }
        /// <summary>Общая трудоёмкость (акад. часы), листы «Курс N» — столбцы 8 и 23.</summary>
        public double? TotalHours { get; set; }
        public double? CreditUnits { get; set; }

        public int? PlanOrder { get; set; }

        public bool RequiresManualValidation { get; set; }

        /// <summary>Строка только из карточки, без пары в плане — выводится отдельным блоком в конце.</summary>
        public bool IsCardOnlyUnmatchedPlan { get; set; }
    }
}
