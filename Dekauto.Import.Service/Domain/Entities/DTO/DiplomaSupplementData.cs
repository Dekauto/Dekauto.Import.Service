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

        /// <summary>Направление подготовки (специальность) из карточки студента.</summary>
        public string? CourseOfTraining { get; set; }

        /// <summary>Целевой объём ОП в з.е. из листа «Свод» (например L8).</summary>
        public double? TargetProgramCredits { get; set; }

        /// <summary>Сумма конт. раб. по блокам 1–3 «ПланСвод» (приоритет для итога ак. час.).</summary>
        public double? TargetContactHoursFromPlan { get; set; }

        /// <summary>Итого з.е. блока 3 (ГИА) из строки-заголовка «ПланСвод».</summary>
        public double? TargetGiaCreditsFromPlan { get; set; }

        /// <summary>Итого з.е. блока 2 «Практика» из строки «ПланСвод» (колонка «Факт»).</summary>
        public double? TargetPracticeCreditsFromPlan { get; set; }

        /// <summary>Наименование ОПОП для листа «4 доп.сведения», ячейка B6 (карточка ~C79).</summary>
        public string? SupplementAdditionalSheetOpopName { get; set; }

        /// <summary>Строка для B7 листа «4 доп.сведения»: «Форма обучения: …» (карточка ~H76).</summary>
        public string? SupplementAdditionalSheetStudyFormLine { get; set; }

        /// <summary>Квалификация для приложения «1 Обладатель диплома» B11 (карточка «ОбщСведения» ~C110).</summary>
        public string? SupplementOwnerQualification { get; set; }

        // Наименования дисциплин (модулей), практик, курсовых работ
        // +Количество зачетных единиц / академических часов / астрономических часов
        // +Оценка
        public List<StudentDisciplineResult> DisciplineResults { get; set; } = new();

    }
}
