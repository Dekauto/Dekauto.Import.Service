using Dekauto.Import.Service.Domain.Entities;
using Dekauto.Import.Service.Domain.Entities.DTO;

namespace Dekauto.Import.Service.Domain.Interfaces
{
    public interface IImportService
    {
        Task<IEnumerable<Student>> GetStudentsLD(IFormFile ld);
        Task<IEnumerable<Student>> GetStudentsContract(IFormFile contract, List<Student> students);
        Task<IEnumerable<Student>> GetStudentsJournal(IFormFile journal, List<Student> students);
        Task<IEnumerable<Student>> GetStudentsStatement(IFormFile statement, List<Student> studentsJournal);
        Task<IEnumerable<Student>> GetStudentsEducationPlan(IFormFile plan, List<Student> students);

        /// <summary>
        /// Отдельный метод парсинга карточки студента для формирования данных для приложения диплома
        /// </summary>
        /// <param name="studentCard">Приходящий по API файл карточки студента для парсинга.</param>
        /// <returns></returns>
        Task<DiplomaSupplementData> GetStudentCardAsync(IFormFile studentCard);
    }
}
