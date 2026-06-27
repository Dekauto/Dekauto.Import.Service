using Dekauto.Import.Service.Domain.Entities;

namespace Dekauto.Import.Service.API.Models
{
    public sealed class StatementImportResult
    {
        public List<Student> Students { get; init; } = new();
        public List<ImportWarning> Warnings { get; init; } = new();
    }
}
