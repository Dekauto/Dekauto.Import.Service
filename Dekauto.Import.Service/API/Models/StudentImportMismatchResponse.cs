namespace Dekauto.Import.Service.API.Models
{
    /// <summary>
    /// Ответ при ошибке согласованности списков студентов между файлами импорта.
    /// </summary>
    public sealed class StudentImportMismatchResponse
    {
        public string ErrorType { get; init; } = "StudentImportMismatch";
        public string Message { get; init; } = string.Empty;
        public string SourceFileName { get; init; } = string.Empty;
        public string SourceFileRole { get; init; } = string.Empty;
        public IReadOnlyList<string> MissingStudents { get; init; } = Array.Empty<string>();
    }
}
