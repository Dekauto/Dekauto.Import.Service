namespace Dekauto.Import.Service.API.Models
{
    public sealed class ImportWarning
    {
        public string Code { get; init; } = string.Empty;
        public string Message { get; init; } = string.Empty;
        public string? FileName { get; init; }
        public string? SheetName { get; init; }
        public string? GroupName { get; init; }
        public string? StudentDisplayName { get; init; }
        public IReadOnlyList<int> MatchedRows { get; init; } = Array.Empty<int>();
    }
}
