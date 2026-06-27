using System.Text;
using System.Text.RegularExpressions;
using Dekauto.Import.Service.Domain.Entities;

namespace Dekauto.Import.Service.Domain.Services
{
    public static class StatementImportMatching
    {
        public const string WarningCodeGroupNotParsed = "StatementGroupNotParsed";
        public const string WarningCodeAmbiguousMatch = "StatementAmbiguousMatch";

        public static string FormatGroupNotParsedUserMessage() =>
            "Не удалось определить номер группы в строке 4. " +
            "Укажите номер группы в строке 4 листа и повторите импорт. Оценки с этого листа не загружены.";

        public static string FormatAmbiguousMatchUserMessage() =>
            "На листе найдено несколько строк с этим ФИО. " +
            "Временно удалите этого студента из файла ведомости и повторите импорт. Оценки этого студента с листа не загружены.";

        private static readonly Regex FullGroupCodePattern = new(
            @"\d{2}\s*[А-ЯA-ZЁ]{2,}(?:[\s\-–—][А-ЯA-ZЁ0-9]+)?(?:\s*\([^)]+\))?(?:\s*[А-ЯA-ZЁ0-9]+(?:[\s\-–—][А-ЯA-ZЁ0-9]+)*)?",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

        private static readonly Regex GroupCodePattern = FullGroupCodePattern;

        private static readonly Regex StatementSurnameRowPattern = new(
            @"^[А-ЯA-ZЁ][а-яa-zё\-]+",
            RegexOptions.CultureInvariant);

        private static readonly int[] StatementGroupSearchRows = { 4 };

        public static string NormalizeStatementFio(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;

            var normalized = value.Replace('\u00A0', ' ').Trim().ToLowerInvariant();
            normalized = normalized.Replace('ё', 'е');
            return normalized.Replace(" ", string.Empty).Replace(".", string.Empty);
        }

        public static bool IsStatementFioHeader(string? header)
        {
            if (string.IsNullOrWhiteSpace(header))
                return false;

            var normalized = header.ToLowerInvariant()
                .Replace('\u00A0', ' ')
                .Replace(".", string.Empty)
                .Replace(" ", string.Empty)
                .Replace("ё", "е");

            if (normalized.Contains("фио", StringComparison.Ordinal))
                return true;

            return normalized.Contains("фамилия", StringComparison.Ordinal)
                   && normalized.Contains("имя", StringComparison.Ordinal);
        }

        public static string NormalizeGroupNameForComparison(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;

            var normalized = value.Replace('\u00A0', ' ').Trim();
            normalized = Regex.Replace(normalized, @"\s+", " ");
            normalized = normalized.ToLowerInvariant();
            normalized = normalized.Replace('–', '-').Replace('—', '-');
            normalized = Regex.Replace(normalized, @"\([бвb]/[оo]\)", string.Empty, RegexOptions.IgnoreCase);
            normalized = Regex.Replace(normalized, @"\([^)]*\)", string.Empty);
            normalized = ApplyGroupHomoglyphs(normalized);
            normalized = normalized.Replace(" ", string.Empty).Replace("-", string.Empty).Replace("/", string.Empty);
            return normalized;
        }

        public static bool GroupNamesMatchForStatement(string? sheetGroupRaw, string? journalGroupName)
        {
            var sheetKey = NormalizeGroupNameForComparison(sheetGroupRaw);
            var journalKey = NormalizeGroupNameForComparison(journalGroupName);
            if (string.IsNullOrEmpty(sheetKey) || string.IsNullOrEmpty(journalKey))
                return false;

            if (sheetKey == journalKey)
                return true;

            return journalKey.StartsWith(sheetKey, StringComparison.Ordinal)
                   || sheetKey.StartsWith(journalKey, StringComparison.Ordinal);
        }

        private static string ApplyGroupHomoglyphs(string value)
        {
            var sb = new StringBuilder(value.Length);
            foreach (var ch in value)
            {
                sb.Append(ch switch
                {
                    'a' => 'а',
                    'c' => 'с',
                    'e' => 'е',
                    'o' => 'о',
                    'p' => 'р',
                    'x' => 'х',
                    'k' => 'к',
                    'm' => 'м',
                    't' => 'т',
                    'b' => 'в',
                    'h' => 'н',
                    'y' => 'у',
                    _ => ch
                });
            }
            return sb.ToString();
        }

        public static bool TryExtractGroupFromStatementRow(
            int row,
            int columnCount,
            Func<int, int, string> getCellText,
            out string groupRaw)
        {
            groupRaw = string.Empty;

            for (int col = 1; col <= columnCount; col++)
            {
                var text = NormalizeStatementSheetCellText(getCellText(row, col));
                if (!ContainsGroupLabel(text))
                    continue;

                if (TryPickGroupCodeFromText(text, preferFullCell: false, out groupRaw))
                    return true;

                foreach (var neighborCol in new[] { col + 1, col - 1, col + 2, col - 2, col + 3 })
                {
                    if (neighborCol < 1 || neighborCol > columnCount)
                        continue;

                    var neighbor = NormalizeStatementSheetCellText(getCellText(row, neighborCol)).Trim();
                    if (string.IsNullOrWhiteSpace(neighbor))
                        continue;
                    if (IsCourseOnlyNeighbor(neighbor))
                        continue;

                    if (TryPickGroupCodeFromText(neighbor, preferFullCell: true, out groupRaw))
                        return true;
                }
            }

            for (int col = 1; col <= columnCount; col++)
            {
                var raw = NormalizeStatementSheetCellText(getCellText(row, col));
                if (string.IsNullOrWhiteSpace(raw) || IsCourseOnlyNeighbor(raw))
                    continue;

                if (TryPickGroupCodeFromText(raw, preferFullCell: true, out groupRaw))
                    return true;
            }

            return false;
        }

        public static bool TryExtractGroupFromStatementSheet(
            int columnCount,
            Func<int, int, string> getCellText,
            out string groupRaw)
        {
            foreach (var row in StatementGroupSearchRows)
            {
                if (TryExtractGroupFromStatementRow(row, columnCount, getCellText, out groupRaw))
                    return true;
            }

            groupRaw = string.Empty;
            return false;
        }

        public static bool TryResolveSheetGroupRaw(
            int columnCount,
            Func<int, int, string> getCellText,
            string? worksheetName,
            IReadOnlyList<Student> students,
            out string groupRaw,
            out bool usedJournalFallback)
        {
            usedJournalFallback = false;
            _ = worksheetName;
            _ = students;
            return TryExtractGroupFromStatementSheet(columnCount, getCellText, out groupRaw);
        }

        public static bool HasProbableStatementStudentRows(
            int columnCount,
            int rowCount,
            Func<int, int, string> getCellText,
            int scanFromRow = 6,
            int maxRowsToScan = 15)
        {
            var rowTo = Math.Min(rowCount, scanFromRow + maxRowsToScan);
            for (int col = 1; col <= columnCount; col++)
            {
                var matches = 0;
                for (int row = scanFromRow; row <= rowTo; row++)
                {
                    var text = ReadStatementCellText(getCellText, row, col);
                    if (string.IsNullOrWhiteSpace(text))
                        continue;

                    if (StatementSurnameRowPattern.IsMatch(text) && text.Contains('.', StringComparison.Ordinal))
                        matches++;
                }

                if (matches >= 2)
                    return true;
            }

            return false;
        }

        public static bool ShouldWarnAboutMissingGroupOnSheet(
            int columnCount,
            int rowCount,
            Func<int, int, string> getCellText,
            bool hasHeaderLayout)
            => hasHeaderLayout || HasProbableStatementStudentRows(columnCount, rowCount, getCellText);

        public static List<Student> FilterStudentsForSheetGroup(IReadOnlyList<Student> students, string sheetGroupRaw)
        {
            var matched = students
                .Where(s => GroupNamesMatchForStatement(sheetGroupRaw, s.GroupName))
                .ToList();
            if (matched.Count > 0)
                return matched;

            return matched;
        }

        public static bool IsStatementGradeSheet(IReadOnlyList<string> headers, Func<string?, bool> isFioHeader)
            => FindFioColumnIndex(headers, isFioHeader) > 0;

        private static bool ContainsGroupLabel(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return false;

            return Regex.IsMatch(text, @"\bгрупп", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }

        private static bool IsCourseOnlyNeighbor(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return false;

            var trimmed = value.Trim();
            if (Regex.IsMatch(trimmed, @"^\d{1,2}$"))
                return true;

            return trimmed.Contains("курс", StringComparison.OrdinalIgnoreCase)
                   && !ContainsGroupLabel(trimmed)
                   && !GroupCodePattern.IsMatch(trimmed);
        }

        private static bool TryPickGroupCodeFromText(string text, bool preferFullCell, out string groupRaw)
        {
            groupRaw = string.Empty;
            if (string.IsNullOrWhiteSpace(text))
                return false;

            var trimmed = NormalizeStatementSheetCellText(text);
            if (preferFullCell && Regex.IsMatch(trimmed, @"^\d{2}", RegexOptions.CultureInvariant))
            {
                var anchored = FullGroupCodePattern.Match(trimmed);
                if (anchored.Success)
                {
                    groupRaw = Regex.Replace(anchored.Value.Trim(), @"\s+", string.Empty);
                    return groupRaw.Length >= 4;
                }
            }

            Match? best = null;
            foreach (Match match in FullGroupCodePattern.Matches(trimmed))
            {
                if (best == null || match.Length > best.Length)
                    best = match;
            }

            if (best == null || !best.Success)
                return false;

            groupRaw = Regex.Replace(best.Value.Trim(), @"\s+", string.Empty);
            return groupRaw.Length >= 4;
        }

        private static string NormalizeStatementSheetCellText(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;
            return Regex.Replace(value.Trim(), @"\s+", " ", RegexOptions.CultureInvariant);
        }

        public static (string KeyFi, string KeyFip) BuildStatementStudentKeys(Student student)
        {
            return BuildStatementNameKeys(student.Surname, student.Name, student.Patronymic);
        }

        public static (string KeyFi, string KeyFip) BuildStatementNameKeys(
            string? surname,
            string? name,
            string? patronymic)
        {
            var surnameValue = surname ?? string.Empty;
            var nameInitial = string.IsNullOrWhiteSpace(name) ? string.Empty : name.Substring(0, 1);
            var patronymicInitial = string.IsNullOrWhiteSpace(patronymic)
                ? string.Empty
                : patronymic.Substring(0, 1);

            var keyFi = NormalizeStatementFio($"{surnameValue}{nameInitial}");
            var keyFip = NormalizeStatementFio($"{surnameValue}{nameInitial}{patronymicInitial}");
            return (keyFi, keyFip);
        }

        public static (string KeyFi, string KeyFip) BuildStatementCellKeys(string? cellRaw)
        {
            if (string.IsNullOrWhiteSpace(cellRaw))
                return (string.Empty, string.Empty);

            var prepared = cellRaw.Replace('\u00A0', ' ').Trim();
            prepared = Regex.Replace(prepared, @"\.+", " ");
            prepared = Regex.Replace(prepared, @"\s+", " ");
            var parts = prepared.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0)
                return (string.Empty, string.Empty);

            var surname = parts[0];
            if (parts.Length == 1)
                return (NormalizeStatementFio(surname), NormalizeStatementFio(surname));

            var second = parts[1];
            string nameInitial;
            string patronymicInitial = string.Empty;

            if (second.Length == 1)
            {
                nameInitial = second;
                if (parts.Length > 2)
                    patronymicInitial = parts[2].Substring(0, 1);
            }
            else if (second.Length == 2 && parts.Length == 2)
            {
                nameInitial = second.Substring(0, 1);
                patronymicInitial = second.Substring(1, 1);
            }
            else
            {
                nameInitial = second.Substring(0, 1);
                if (parts.Length > 2)
                    patronymicInitial = parts[2].Substring(0, 1);
            }

            return BuildStatementNameKeys(surname, nameInitial, patronymicInitial);
        }

        public static bool StatementFioCellMatchesStudent(string? cellRaw, string keyFi, string keyFip)
        {
            if (string.IsNullOrWhiteSpace(cellRaw))
                return false;

            var compact = NormalizeStatementFio(cellRaw);
            if (!string.IsNullOrEmpty(compact))
            {
                if (!string.IsNullOrEmpty(keyFip) && compact == keyFip)
                    return true;

                if (!string.IsNullOrEmpty(keyFi) && compact == keyFi)
                    return true;
            }

            var (cellFi, cellFip) = BuildStatementCellKeys(cellRaw);
            if (string.IsNullOrEmpty(cellFi) && string.IsNullOrEmpty(cellFip))
                return false;

            if (!string.IsNullOrEmpty(keyFip) && cellFip == keyFip)
                return true;

            if (!string.IsNullOrEmpty(keyFi) && cellFi == keyFi)
            {
                if (!string.IsNullOrEmpty(keyFip) && keyFip != keyFi && !string.IsNullOrEmpty(cellFip) && cellFip != keyFi)
                    return false;
                return true;
            }

            return false;
        }

        public static string ReadStatementCellText(Func<int, int, string> getCellText, int row, int col)
        {
            var text = getCellText(row, col);
            return string.IsNullOrWhiteSpace(text) ? string.Empty : text.Trim();
        }

        public static List<int> FindStatementFioCandidateRows(
            Func<int, int, string> getCellText,
            int fioCol,
            int rowFrom,
            int rowTo,
            string keyFi,
            string keyFip)
        {
            var candidates = new List<int>();
            if (fioCol < 1)
                return candidates;

            for (int row = rowFrom; row <= rowTo; row++)
            {
                var cellRaw = ReadStatementCellText(getCellText, row, fioCol);
                if (StatementFioCellMatchesStudent(cellRaw, keyFi, keyFip))
                    candidates.Add(row);
            }

            return candidates;
        }

        public static int FindFioColumnIndex(IReadOnlyList<string> headers, Func<string?, bool> isFioHeader)
        {
            for (int col = 0; col < headers.Count; col++)
            {
                if (isFioHeader(headers[col]))
                    return col + 1;
            }
            return -1;
        }

        public static int FindFioColumnIndex(
            IReadOnlyList<string> headers,
            IReadOnlyList<string> topHeaders,
            IReadOnlyList<string> bottomHeaders,
            Func<string?, bool> isFioHeader)
        {
            var columnCount = Math.Max(headers.Count, Math.Max(topHeaders.Count, bottomHeaders.Count));
            for (int col = 0; col < columnCount; col++)
            {
                if (col < headers.Count && isFioHeader(headers[col]))
                    return col + 1;
                if (col < topHeaders.Count && isFioHeader(topHeaders[col]))
                    return col + 1;
                if (col < bottomHeaders.Count && isFioHeader(bottomHeaders[col]))
                    return col + 1;
            }

            return -1;
        }

        public static bool TryFindStatementHeaderLayout(
            int columnCount,
            Func<int, int, string> getCellText,
            Func<string?, string> normalizeHeader,
            Func<string?, bool> isFioHeader,
            out int headerTopRow,
            out int headerBottomRow,
            out List<string> headers,
            out List<string> topHeaders,
            out List<string> bottomHeaders)
        {
            headers = new List<string>();
            topHeaders = new List<string>();
            bottomHeaders = new List<string>();
            headerTopRow = 6;
            headerBottomRow = 7;

            foreach (var (topRow, bottomRow) in new[] { (6, 7), (7, 8), (5, 6), (8, 9), (6, 8) })
            {
                var candidateHeaders = new List<string>();
                var candidateTopHeaders = new List<string>();
                var candidateBottomHeaders = new List<string>();

                for (int col = 1; col <= columnCount; col++)
                {
                    var top = normalizeHeader(getCellText(topRow, col));
                    var bottom = normalizeHeader(getCellText(bottomRow, col));
                    candidateTopHeaders.Add(top);
                    candidateBottomHeaders.Add(bottom);

                    string combined;
                    if (string.IsNullOrWhiteSpace(top))
                        combined = bottom;
                    else if (string.IsNullOrWhiteSpace(bottom))
                        combined = top;
                    else
                        combined = $"{top} {bottom}";

                    candidateHeaders.Add(normalizeHeader(combined));
                }

                if (FindFioColumnIndex(candidateHeaders, candidateTopHeaders, candidateBottomHeaders, isFioHeader) < 1)
                    continue;

                headers = candidateHeaders;
                topHeaders = candidateTopHeaders;
                bottomHeaders = candidateBottomHeaders;
                headerTopRow = topRow;
                headerBottomRow = bottomRow;
                return true;
            }

            return false;
        }

        public static int FindStatementStudentRowStart(
            Func<int, int, string> getCellText,
            int fioCol,
            int searchFrom,
            int rowTo)
        {
            for (int row = searchFrom; row <= Math.Min(searchFrom + 8, rowTo); row++)
            {
                var text = ReadStatementCellText(getCellText, row, fioCol);
                if (string.IsNullOrWhiteSpace(text))
                    continue;

                if (StatementSurnameRowPattern.IsMatch(text))
                    return row;
            }

            return searchFrom;
        }

        public static int GetStatementStudentRowStart(int headerBottomRow) => headerBottomRow + 1;

        public static string FormatStudentDisplayName(Student student)
        {
            var parts = new[] { student.Surname, student.Name, student.Patronymic }
                .Where(p => !string.IsNullOrWhiteSpace(p));
            var text = string.Join(' ', parts);
            return string.IsNullOrWhiteSpace(text) ? "(без ФИО)" : text;
        }
    }
}
