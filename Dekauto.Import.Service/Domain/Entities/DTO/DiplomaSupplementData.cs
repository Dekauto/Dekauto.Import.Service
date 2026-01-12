namespace Dekauto.Import.Service.Domain.Entities.DTO
{
    /// <summary>
    /// Модель для наполнения данными, требуемыми при заполнении приложения диплома в экспорте. Парсится здесь, из карточки студента.
    /// </summary>
    public class DiplomaSupplementData
    {
        public bool? DiplomaWithHonors { get; set; } // с отличием
        public string? Name { get; set; } // Имя
        public string? Surname { get; set; } // Фамилия
        public string? Patronymic { get; set; } // Отчество
        public DateOnly? BirthdayDate { get; set; } // Дата рождения
        public string? EducationReceived { get; set; } // Наименование документа о предыдущем образовании
        public DateOnly? EducationReceivedDate { get; set; } // Год выдачи документа образования

        // Наименования дисциплин (модулей), практик, курсовых работ
        // +Количество зачетных единиц / академических часов / астрономических часов
        // +Оценка
        public List<StudentDisciplineResult> DisciplineResults { get; set; } = new();

    }
}
