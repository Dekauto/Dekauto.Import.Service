namespace Dekauto.Import.Service.Domain.Exceptions
{
    /// <summary>
    /// Обучающиеся из предыдущих шагов импорта не найдены в журнале зачётных книжек (несовпадение при сопоставлении).
    /// </summary>
    public sealed class StudentImportMismatchException : Exception
    {
        public string SourceFileRole { get; }
        public string SourceFileName { get; }
        public IReadOnlyList<string> MissingStudents { get; }

        public StudentImportMismatchException(
            string sourceFileRole,
            string sourceFileName,
            IReadOnlyList<string> missingStudents)
            : base(BuildMessage(sourceFileRole, sourceFileName, missingStudents))
        {
            SourceFileRole = sourceFileRole;
            SourceFileName = sourceFileName;
            MissingStudents = missingStudents;
        }

        private static string BuildMessage(string role, string fileName, IReadOnlyList<string> missing)
        {
            var list = string.Join("; ", missing);
            return $"В файле «{fileName}» ({role}) отсутствуют строки, соответствующие следующим обучающимся: {list}. " +
                "Для устранения несоответствия удалите сведения о перечисленных обучающихся из первого файла импорта (личные дела) " +
                "либо добавьте в третий файл импорта (журнал учёта выдачи зачётных книжек) строки с корректными данными о группах для каждого из них.";
        }
    }
}
