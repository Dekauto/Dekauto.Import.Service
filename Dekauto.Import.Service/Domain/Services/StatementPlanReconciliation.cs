using System.Text.RegularExpressions;
using Dekauto.Import.Service.Domain.Entities;
using Dekauto.Import.Service.Domain.Entities.DTO;

namespace Dekauto.Import.Service.Domain.Services
{
    public static class StatementPlanReconciliation
    {
        public const double PlanDisciplineFuzzyThreshold = 0.9;
        public const double PlanPracticeFuzzyThreshold = 0.6;
        private const int PlanPracticePrefixMinLength = 20;

        public static bool IsStatementPracticeDiscipline(string? disciplineName)
        {
            if (string.IsNullOrWhiteSpace(disciplineName))
                return false;

            if (Regex.IsMatch(disciplineName, @"^Учебная практика,", RegexOptions.IgnoreCase))
                return true;
            if (Regex.IsMatch(disciplineName, @"^Производственная практика,", RegexOptions.IgnoreCase))
                return true;
            if (Regex.IsMatch(disciplineName, @"^Преддипломная практика", RegexOptions.IgnoreCase))
                return true;
            if (Regex.IsMatch(disciplineName, @"^Практика[\s,]", RegexOptions.IgnoreCase))
                return true;

            return false;
        }

        public static bool IsPracticeCategoryHeader(string? disciplineName)
        {
            var n = NormalizeDisciplineName(disciplineName);
            if (string.IsNullOrWhiteSpace(n))
                return false;

            return Regex.IsMatch(
                n,
                @"^(Учебная|Производственная|Преддипломная)\s+практика\s*$",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }

        public static string FormatPracticeDisplayName(string typeTitle, string itemTitle)
        {
            typeTitle = NormalizeDisciplineName(typeTitle);
            itemTitle = NormalizeDisciplineName(itemTitle);
            if (string.IsNullOrWhiteSpace(typeTitle))
                return itemTitle;
            if (string.IsNullOrWhiteSpace(itemTitle))
                return typeTitle;

            static string CapFirst(string s)
            {
                s = s.Trim();
                if (s.Length == 0) return s;
                var c = s[0];
                return char.IsUpper(c) ? s : char.ToUpperInvariant(c) + s.Substring(1);
            }

            static string LowerFirst(string s)
            {
                s = s.Trim();
                if (s.Length == 0) return s;
                var c = s[0];
                return char.IsLower(c) ? s : char.ToLowerInvariant(c) + s.Substring(1);
            }

            return CapFirst(typeTitle) + ", " + LowerFirst(itemTitle);
        }

        public static void UpsertSnapshot(List<StatementDisciplineSnapshot> list, StatementDisciplineSnapshot snapshot)
        {
            var idx = list.FindIndex(s => s.MergeKey == snapshot.MergeKey);
            if (idx >= 0)
                list[idx] = snapshot;
            else
                list.Add(snapshot);
        }

        public static void ReconcileStudentsWithPlan(IEnumerable<Student> students, IReadOnlyList<PlanDisciplineCatalogEntry> catalog)
        {
            foreach (var student in students)
            {
                var results = new List<StudentDisciplineResult>();
                foreach (var snapshot in student.PendingStatementDisciplines)
                {
                    results.Add(ReconcileSnapshot(snapshot, catalog));
                }

                student.DisciplineResults = results;
                student.PendingStatementDisciplines.Clear();
            }
        }

        public static StudentDisciplineResult ReconcileSnapshot(
            StatementDisciplineSnapshot snapshot,
            IReadOnlyList<PlanDisciplineCatalogEntry> catalog)
        {
            if (snapshot.IsCourseWork)
            {
                return new StudentDisciplineResult
                {
                    DisciplineName = null,
                    Score = snapshot.Score,
                    Semester = snapshot.Semester,
                    Year = snapshot.Year,
                    ControlType = "курсовая"
                };
            }

            if (!snapshot.Semester.HasValue)
            {
                return BuildStatementOnlyResult(snapshot, requiresReview: true);
            }

            var semester = snapshot.Semester.Value;
            var semesterCatalog = catalog.Where(c => c.Semester == semester).ToList();

            if (snapshot.IsPractice)
            {
                var practiceMatch = FindBestPracticeMatch(snapshot.StatementName, semesterCatalog);
                if (practiceMatch != null)
                {
                    return BuildFromPlanEntry(snapshot, practiceMatch, requiresReview: true);
                }

                return BuildStatementOnlyResult(snapshot, requiresReview: true);
            }

            // Обычные дисциплины сопоставляются по PlanName независимо от IsPractice в каталоге:
            // после блока практик надзаголовок может оставаться, но название строки плана — своё.
            var exact = semesterCatalog.FirstOrDefault(c =>
                LetterExactMatch(snapshot.StatementName, c.PlanName));
            if (exact != null)
            {
                return BuildFromPlanEntry(snapshot, exact, requiresReview: false);
            }

            var fuzzy = FindBestRegularFuzzyMatch(snapshot.StatementName, semesterCatalog, PlanDisciplineFuzzyThreshold);
            if (fuzzy != null)
            {
                return BuildFromPlanEntry(snapshot, fuzzy, requiresReview: true);
            }

            return BuildStatementOnlyResult(snapshot, requiresReview: true);
        }

        public static bool LetterExactMatch(string? statementName, string? planName)
        {
            var a = NormalizeDisciplineName(statementName);
            var b = NormalizeDisciplineName(planName);
            if (a.Length == 0 || b.Length == 0)
                return false;
            return string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
        }

        public static double DisciplineNameSimilarityRatio(string? a, string? b)
        {
            var na = NormalizeDisciplineName(a).ToLowerInvariant();
            var nb = NormalizeDisciplineName(b).ToLowerInvariant();
            if (na.Length == 0 && nb.Length == 0) return 1d;
            if (na.Length == 0 || nb.Length == 0) return 0d;
            var dist = LevenshteinDistance(na, nb);
            var mx = Math.Max(na.Length, nb.Length);
            return 1d - (double)dist / mx;
        }

        private static PlanDisciplineCatalogEntry? FindBestRegularFuzzyMatch(
            string statementName,
            IEnumerable<PlanDisciplineCatalogEntry> catalog,
            double threshold)
        {
            PlanDisciplineCatalogEntry? best = null;
            var bestRatio = 0d;
            foreach (var entry in catalog)
            {
                var ratio = DisciplineNameSimilarityRatio(statementName, entry.PlanName);
                if (ratio >= threshold && ratio > bestRatio)
                {
                    bestRatio = ratio;
                    best = entry;
                }
            }

            return best;
        }

        private static PlanDisciplineCatalogEntry? FindBestPracticeMatch(
            string statementName,
            IEnumerable<PlanDisciplineCatalogEntry> catalog)
        {
            PlanDisciplineCatalogEntry? best = null;
            var bestScore = 0d;
            foreach (var entry in catalog.Where(c => c.IsPractice))
            {
                var score = PracticePlanNameScore(entry.DisplayName, statementName);
                if (score >= PlanPracticeFuzzyThreshold && score > bestScore)
                {
                    bestScore = score;
                    best = entry;
                }
            }

            return best;
        }

        private static double PracticePlanNameScore(string planComposite, string? statementName)
        {
            if (string.IsNullOrWhiteSpace(statementName))
                return 0d;

            var a = NormalizeDisciplineName(planComposite).ToLowerInvariant();
            var b = NormalizeDisciplineName(statementName).ToLowerInvariant();
            if (a.Length == 0 || b.Length == 0)
                return 0d;

            var shorter = a.Length <= b.Length ? a : b;
            var longer = a.Length <= b.Length ? b : a;
            if (shorter.Length >= PlanPracticePrefixMinLength && longer.StartsWith(shorter, StringComparison.Ordinal))
                return 1d;

            return DisciplineNameSimilarityRatio(planComposite, statementName);
        }

        private static StudentDisciplineResult BuildFromPlanEntry(
            StatementDisciplineSnapshot snapshot,
            PlanDisciplineCatalogEntry entry,
            bool requiresReview)
        {
            var isPracticeResult = snapshot.IsPractice;
            return new StudentDisciplineResult
            {
                DisciplineName = isPracticeResult ? entry.DisplayName : entry.PlanName,
                Score = snapshot.Score,
                Semester = snapshot.Semester,
                Year = snapshot.Year,
                ControlType = isPracticeResult ? "практика" : entry.ControlType,
                TotalHours = entry.TotalHours,
                AudHours = entry.AudHours,
                RequiresManualValidation = requiresReview
            };
        }

        private static StudentDisciplineResult BuildStatementOnlyResult(
            StatementDisciplineSnapshot snapshot,
            bool requiresReview)
        {
            return new StudentDisciplineResult
            {
                DisciplineName = string.IsNullOrWhiteSpace(snapshot.StatementName) ? null : snapshot.StatementName,
                Score = snapshot.Score,
                Semester = snapshot.Semester,
                Year = snapshot.Year,
                ControlType = snapshot.IsPractice ? "практика" : null,
                RequiresManualValidation = requiresReview
            };
        }

        private static string NormalizeDisciplineName(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;
            return Regex.Replace(value.Trim(), @"\s+", " ", RegexOptions.CultureInvariant);
        }

        private static int LevenshteinDistance(string a, string b)
        {
            var n = a.Length;
            var m = b.Length;
            if (n == 0) return m;
            if (m == 0) return n;
            var d = new int[n + 1, m + 1];
            for (var i = 0; i <= n; i++) d[i, 0] = i;
            for (var j = 0; j <= m; j++) d[0, j] = j;
            for (var i = 1; i <= n; i++)
            {
                for (var j = 1; j <= m; j++)
                {
                    var cost = char.ToLowerInvariant(a[i - 1]) == char.ToLowerInvariant(b[j - 1]) ? 0 : 1;
                    d[i, j] = Math.Min(Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1), d[i - 1, j - 1] + cost);
                }
            }

            return d[n, m];
        }
    }
}
