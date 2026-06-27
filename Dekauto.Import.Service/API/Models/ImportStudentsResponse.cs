using Dekauto.Import.Service.Domain.Entities;

namespace Dekauto.Import.Service.API.Models
{
    public sealed class ImportStudentsResponse
    {
        public List<Student> Students { get; init; } = new();
        public List<ImportWarning> ImportWarnings { get; init; } = new();
    }
}
