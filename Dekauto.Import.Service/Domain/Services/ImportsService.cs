using Dekauto.Import.Service.Domain.Entities;
using Dekauto.Import.Service.Domain.Entities.DTO;
using Dekauto.Import.Service.Domain.Exceptions;
using Dekauto.Import.Service.Domain.Interfaces;
using OfficeOpenXml;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace Dekauto.Import.Service.Domain.Services
{
    public class ImportsService : IImportService
    {
        private IConfiguration configuration;
        private readonly ILogger<ImportsService> logger;
        public ImportsService(IConfiguration configuration, ILogger<ImportsService> logger)
        {
            this.configuration = configuration;
            this.logger = logger;
        }

        private static string FormatStudentDisplayName(Student student)
        {
            var s = $"{student.Surname} {student.Name} {student.Patronymic}".Trim();
            return string.IsNullOrWhiteSpace(s) ? "(пустое ФИО)" : s;
        }

        /// <summary>
        /// Номер дома: только после маркера «д.» / «дом» и цифры (не путать с «д.» в типе нас. пункта и не брать длинные слова).
        /// </summary>
        private static string? ParseAddressHouseNumber(string address)
        {
            if (string.IsNullOrWhiteSpace(address))
                return null;

            var matches = Regex.Matches(
                address,
                @"(?:^|[,\s])(?:дом|д)\s*\.?\s*(\d[\w\-/]*)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

            if (matches.Count == 0)
                return null;

            var house = matches[matches.Count - 1].Groups[1].Value.Trim();
            if (house.Length > 10)
                house = house[..10];
            return string.IsNullOrEmpty(house) ? null : house;
        }

        /// <summary>
        /// Тип населённого пункта (г/с/х/д/п). Маркер «д. 15» (дом) не считается типом «деревня».
        /// </summary>
        private static string? ParseAddressSettlementTypeAbbrev(string address)
        {
            if (string.IsNullOrWhiteSpace(address))
                return null;

            var afterName = Regex.Match(
                address,
                @"\b([\p{L}\s\-]+?)\s+([гсхдп])(?:\.|\s|,|$)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (afterName.Success)
                return afterName.Groups[2].Value.ToLowerInvariant();

            var beforeName = Regex.Match(
                address,
                @"\b([гсхдп])\.(?!\s*\d)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (beforeName.Success)
                return beforeName.Groups[1].Value.ToLowerInvariant();

            return null;
        }

        private static void ApplyAddressSettlementType(Student student, string? abbr, bool isRegistration)
        {
            if (string.IsNullOrEmpty(abbr))
                return;

            var type = abbr switch
            {
                "г" => "город",
                "с" => "село",
                "х" => "хутор",
                "д" => "деревня",
                "п" => "посёлок",
                _ => null
            };

            if (type == null)
                return;

            if (isRegistration)
                student.AddressRegistrationType = type;
            else
                student.AddressResidentialType = type;
        }

        private static string NormalizeStatementSheetCellText(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;
            return Regex.Replace(value.Trim(), @"\s+", " ", RegexOptions.CultureInvariant);
        }

        private static bool ContainsStatementDatePattern(string text) =>
            Regex.IsMatch(text, @"\d{1,2}\.\d{1,2}\.\d{2,4}", RegexOptions.CultureInvariant);

        /// <summary>
        /// Явное указание курса в одной ячейке: «3 курс», «курс 3».
        /// </summary>
        private static bool TryExtractExplicitStatementCourseNumber(string? text, out int courseNum)
        {
            courseNum = default;
            if (string.IsNullOrWhiteSpace(text))
                return false;

            var explicitCourse = Regex.Match(
                text,
                @"(?:(\d)\s*[-–]?\s*курс|курс\s*(\d))",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (!explicitCourse.Success)
                return false;

            var digits = explicitCourse.Groups[1].Success && !string.IsNullOrEmpty(explicitCourse.Groups[1].Value)
                ? explicitCourse.Groups[1].Value
                : explicitCourse.Groups[2].Value;
            return int.TryParse(digits, out courseNum) && courseNum >= 1 && courseNum <= 6;
        }

        /// <summary>
        /// Номер курса из текста ячейки: не брать части даты (02.02.2023) и годы.
        /// </summary>
        private static bool TryExtractStatementCourseNumber(string? text, out int courseNum)
        {
            courseNum = default;
            if (string.IsNullOrWhiteSpace(text))
                return false;

            if (TryExtractExplicitStatementCourseNumber(text, out courseNum))
                return true;

            // В ячейках с датой документа не угадываем курс по первому числу (02 → 2-й курс)
            if (ContainsStatementDatePattern(text))
                return false;

            foreach (Match numberMatch in Regex.Matches(text, @"\d+"))
            {
                var idx = numberMatch.Index;
                var len = numberMatch.Length;
                if (idx > 0 && text[idx - 1] == '.')
                    continue;
                if (idx + len < text.Length && text[idx + len] == '.')
                    continue;

                if (numberMatch.Value.Length == 4
                    && int.TryParse(numberMatch.Value, out var year)
                    && year >= 1900
                    && year <= 2100)
                    continue;

                if (int.TryParse(numberMatch.Value, out courseNum) && courseNum >= 1 && courseNum <= 6)
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Строка 4 ведомости: «курс» и цифра часто в соседних ячейках; шапка с датой может быть в объединённой ячейке.
        /// </summary>
        private static bool TryExtractCourseFromStatementRow(
            int row,
            int columnCount,
            Func<int, int, string> getCellText,
            out int courseNum)
        {
            courseNum = default;

            for (int col = 1; col <= columnCount; col++)
            {
                var text = NormalizeStatementSheetCellText(getCellText(row, col));
                if (!text.Contains("курс", StringComparison.OrdinalIgnoreCase))
                    continue;

                if (TryExtractExplicitStatementCourseNumber(text, out courseNum))
                    return true;

                foreach (var neighborCol in new[] { col + 1, col - 1, col + 2, col - 2 })
                {
                    if (neighborCol < 1 || neighborCol > columnCount)
                        continue;

                    var neighbor = NormalizeStatementSheetCellText(getCellText(row, neighborCol)).Trim();
                    if (Regex.IsMatch(neighbor, @"^\d{1,2}$", RegexOptions.CultureInvariant)
                        && int.TryParse(neighbor, out courseNum)
                        && courseNum >= 1
                        && courseNum <= 6)
                        return true;
                }
            }

            for (int col = 1; col <= columnCount; col++)
            {
                var raw = getCellText(row, col);
                if (string.IsNullOrWhiteSpace(raw))
                    continue;

                if (TryExtractStatementCourseNumber(raw, out courseNum))
                    return true;
            }

            return false;
        }

        private static string NormalizeDisciplineName(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;
            return Regex.Replace(value, @"\s+", " ").Trim();
        }

        private static string NormalizeDisciplineNameForComparison(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;

            var normalized = value.Trim();
            var studyMatch = Regex.Match(normalized, @"^Учебная практика,\s*", RegexOptions.IgnoreCase);
            if (studyMatch.Success)
            {
                normalized = normalized.Substring(studyMatch.Length).Trim();
            }
            else
            {
                var productionMatch = Regex.Match(normalized, @"^Производственная практика,\s*", RegexOptions.IgnoreCase);
                if (productionMatch.Success)
                {
                    normalized = normalized.Substring(productionMatch.Length).Trim();
                }
            }

            var practicaComma = Regex.Match(normalized, @"^(.+?практика),\s*", RegexOptions.IgnoreCase);
            if (practicaComma.Success)
            {
                normalized = normalized.Substring(practicaComma.Length).Trim();
            }

            normalized = Regex.Replace(normalized, @"\bпрофессионально\s*$",
                "профессиональной деятельности",
                RegexOptions.IgnoreCase).Trim();

            normalized = Regex.Replace(normalized, @"профессионально\s+деятельности",
                    "профессиональной деятельности",
                    RegexOptions.IgnoreCase)
                .Trim();

            return normalized;
        }

        /// <summary>Первая строка ячейки кода индекса (в планах часто многострочный текст с переносами).</summary>
        private static string SupplementPlanCodeCellFirstLine(string? raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
                return string.Empty;
            var t = raw.Replace('\u00A0', ' ').Trim();
            var nl = t.IndexOfAny(new[] { '\r', '\n' });
            if (nl >= 0)
                t = t.Substring(0, nl).Trim();
            return t;
        }

        private static string StripSyntheticCourseTopicFromCard(string? cardName)
        {
            if (string.IsNullOrWhiteSpace(cardName))
                return string.Empty;
            var s = NormalizeDisciplineName(cardName);
            const string prefix = "НАЗВАНИЕ ДИСЦИПЛИНЫ \"";
            if (s.Length > prefix.Length + 1 &&
                s.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
                s.EndsWith("\"", StringComparison.Ordinal))
            {
                return s.Substring(prefix.Length, s.Length - prefix.Length - 1).Trim();
            }
            return s;
        }

        private static bool CardCourseWorkLooksLikePlaceholderTopic(StudentDisciplineResult c)
        {
            var topic = StripSyntheticCourseTopicFromCard(c.DisciplineName);
            return string.IsNullOrWhiteSpace(topic) ||
                   string.Equals(topic, "Тема курсовой работы", StringComparison.OrdinalIgnoreCase);
        }

        private static bool PlanNameLooksLikeCourseWorkPlanRow(string? planName)
        {
            if (string.IsNullOrWhiteSpace(planName))
                return false;
            var n = NormalizeDisciplineName(planName).ToLowerInvariant();
            if (n.Contains("курсовая работа"))
                return true;
            if (n.Contains("курсовой проект"))
                return true;
            return false;
        }

        private bool IsSupplementPlanIncludedByInPlanMarker(string? rawMark)
        {
            var m = rawMark?.Replace('\u00A0', ' ').Trim() ?? "";

            if (string.Equals(m, "+", StringComparison.OrdinalIgnoreCase))
                return true;

            if (configuration.GetValue("Import:SupplementStudyPlan:RequireExplicitPlusInPlan", false))
                return false;

            if (string.IsNullOrWhiteSpace(m))
                return configuration.GetValue("Import:SupplementStudyPlan:TreatEmptyInPlanMarkerAsIncluded", true);

            if (m == "-" || m == "–" || m == "—")
                return false;
            if (string.Equals(m, "нет", StringComparison.OrdinalIgnoreCase))
                return false;

            var lower = m.ToLowerInvariant();
            if (lower == "да" || lower == "1" || lower == "v")
                return true;
            if (m.Contains('✓') || m.Contains('✔') || m.Contains('☑'))
                return true;

            return false;
        }

        private static double DisciplineNameSimilarityRatioForMatch(string? planName, string? cardName)
        {
            var direct = DisciplineNameSimilarityRatio(planName, cardName);
            var topic = StripSyntheticCourseTopicFromCard(cardName);
            if (string.IsNullOrWhiteSpace(topic))
                return direct;
            var topicNorm = NormalizeDisciplineName(topic);
            var cardNorm = NormalizeDisciplineName(cardName);
            if (string.Equals(topicNorm, cardNorm, StringComparison.OrdinalIgnoreCase))
                return direct;
            var viaTopic = DisciplineNameSimilarityRatio(planName, topic);
            return direct > viaTopic ? direct : viaTopic;
        }

        private static bool ExactPlanCardNameMatch(string planName, string? cardName)
        {
            if (string.IsNullOrWhiteSpace(cardName))
                return false;
            var pn = NormalizeDisciplineName(planName);
            var cn = NormalizeDisciplineName(cardName);
            if (string.Equals(pn, cn, StringComparison.OrdinalIgnoreCase))
                return true;
            var topic = StripSyntheticCourseTopicFromCard(cardName);
            return !string.IsNullOrWhiteSpace(topic) &&
                   string.Equals(pn, NormalizeDisciplineName(topic), StringComparison.OrdinalIgnoreCase);
        }

        private static bool RelaxedPlanCardNameMatch(string planName, string? cardName)
        {
            if (string.IsNullOrWhiteSpace(cardName))
                return false;
            var a = NormalizeDisciplineNameForComparison(planName);
            var b = NormalizeDisciplineNameForComparison(cardName);
            if (string.Equals(a, b, StringComparison.OrdinalIgnoreCase))
                return true;
            var topic = StripSyntheticCourseTopicFromCard(cardName);
            return !string.IsNullOrWhiteSpace(topic) &&
                   string.Equals(a, NormalizeDisciplineNameForComparison(topic), StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsTruncatedPlanCardNamePrefixMatch(string planName, string? cardName, int minShorterLen = 30)
        {
            if (string.IsNullOrWhiteSpace(planName) || string.IsNullOrWhiteSpace(cardName))
                return false;

            var a = NormalizeDisciplineNameForComparison(planName).Trim();
            var b = NormalizeDisciplineNameForComparison(cardName).Trim();
            if (a.Length == 0 || b.Length == 0)
                return false;

            var shorter = a.Length <= b.Length ? a : b;
            var longer = a.Length <= b.Length ? b : a;
            if (shorter.Length < minShorterLen)
                return false;

            return longer.StartsWith(shorter, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>«Тип практики» + строка названия как в приложении к диплому.</summary>
        private static string FormatPracticePlanNameWithTypePrefix(string typeTitle, string itemTitle)
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

        private static bool LooksLikePlanGiaDefensePreparationExcluded(string normalizedName)
        {
            var lower = normalizedName.ToLowerInvariant();
            return lower.Contains("подготовк", StringComparison.Ordinal)
                   && lower.Contains("процедур", StringComparison.Ordinal)
                   && lower.Contains("защит", StringComparison.Ordinal);
        }

        private static bool LoosePlanPracticeNameEquals(string planName, string? cardName)
        {
            if (string.IsNullOrWhiteSpace(cardName))
                return false;
            var pa = NormalizeDisciplineName(planName).ToLowerInvariant();
            var pb = NormalizeDisciplineName(cardName).ToLowerInvariant();
            if (!pa.Contains("практик") || !pb.Contains("практик"))
                return false;
            if (!pa.Contains("профессиональ") && !pb.Contains("профессиональ"))
                return false;

            var a = NormalizeDisciplineNameForComparison(planName);
            var b = NormalizeDisciplineNameForComparison(cardName);
            return string.Equals(a.Trim(), b.Trim(), StringComparison.OrdinalIgnoreCase);
        }

        private const double PlanCardFuzzySimilarityThreshold = 0.82;

        private static int LevenshteinDistance(string a, string b)
        {
            int n = a.Length;
            int m = b.Length;
            if (n == 0) return m;
            if (m == 0) return n;
            var d = new int[n + 1, m + 1];
            for (int i = 0; i <= n; i++) d[i, 0] = i;
            for (int j = 0; j <= m; j++) d[0, j] = j;
            for (int i = 1; i <= n; i++)
            {
                for (int j = 1; j <= m; j++)
                {
                    int cost = char.ToLowerInvariant(a[i - 1]) == char.ToLowerInvariant(b[j - 1]) ? 0 : 1;
                    d[i, j] = Math.Min(Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1), d[i - 1, j - 1] + cost);
                }
            }
            return d[n, m];
        }

        private static double DisciplineNameSimilarityRatio(string? planName, string? cardName)
        {
            var na = NormalizeDisciplineName(planName).ToLowerInvariant();
            var nb = NormalizeDisciplineName(cardName).ToLowerInvariant();
            if (na.Length == 0 && nb.Length == 0) return 1d;
            if (na.Length == 0 || nb.Length == 0) return 0d;
            int dist = LevenshteinDistance(na, nb);
            int mx = Math.Max(na.Length, nb.Length);
            return 1d - (double)dist / mx;
        }

        public async Task<IEnumerable<Student>> GetStudentsContract(IFormFile contract, List<Student> students)
        {
            using (var stream = new MemoryStream())
            {
                await contract.CopyToAsync(stream);
                using (var package = new ExcelPackage(stream))
                {
                    var worksheet = package.Workbook.Worksheets[0] ?? throw new InvalidOperationException("Загруженный файл не содержит листов");

                    var columnCount = worksheet.Dimension.Columns;
                    var rowCount = worksheet.Dimension.Rows;

                    var headers = new List<string>();
                    for (int col = 1; col <= columnCount; col++)
                    {
                        headers.Add(worksheet.Cells[1, col].Text);
                    }
                    if (students.Count == 0 || students == null) throw new ArgumentNullException("Студенты отсутствуют в таблице личных дел");
                    foreach (var student in students)
                    {
                        string fio = $"{student.Surname}{student.Name}{student.Patronymic}".ToLower();
                        for (int row = 2; row <= rowCount; row++)
                        {
                            bool isCurrentStudent = false;
                            string enrollementOrderDatePattern = @"\d{2}\.\d{2}\.\d{4}";
                            string enrollementOrderNumPattern = @"\d*\-\d*\/\d*";
                            for (int col = 1; col <= columnCount; col++)
                            {
                                var header = headers[col - 1];
                                var cellValue = worksheet.Cells[row, col].Value ?? "";

                                if (header.ToLower() == "фио обучающегося" || header.ToLower() == "фио студента")
                                {
                                    string cellfio = cellValue.ToString().ToLower().Replace(" ", "");
                                    if (cellfio == fio)
                                    {
                                        isCurrentStudent = true;
                                        break;
                                    }
                                }
                            }
                            for (int col = 1; col <= columnCount; col++)
                            {
                                var header = headers[col - 1];
                                var cellValue = worksheet.Cells[row, col].Value ?? "";

                                logger.LogInformation($"Работа с ячейкой: [{col},{row}]; столбец {header}");

                                switch (header.ToLower())
                                {
                                    case "дата":
                                        if (isCurrentStudent == true)
                                        {
                                            if (cellValue is DateTime excelDate)
                                            {
                                                student.EducationRelationDate = DateOnly.FromDateTime(excelDate);
                                            }
                                            else
                                            {
                                                string dateStr = cellValue.ToString().Trim();
                                                if (DateTime.TryParseExact(
                                                    dateStr,
                                                    "dd.MM.yyyy",
                                                    CultureInfo.InvariantCulture,
                                                    DateTimeStyles.None,
                                                    out DateTime parsedDate))
                                                {
                                                    student.EducationRelationDate = DateOnly.FromDateTime(parsedDate);
                                                }
                                                else throw new FormatException($"Не удалось распознать дату: {dateStr}");
                                            }
                                            student.EducationStartYear = short.Parse(student.EducationRelationDate.Value.Year.ToString());
                                            student.EducationFinishYear = (short)(student.EducationStartYear + student.EducationTime);
                                        }
                                        break;
                                    case "№ договора":
                                    case "номер договора":
                                        if (isCurrentStudent == true)
                                        {
                                            student.EducationRelationNum = cellValue.ToString();
                                        }
                                        break;
                                    case "№ приказа о зачислении":
                                    case "номер приказа о зачислении":
                                        if (isCurrentStudent == true)
                                        {
                                            string dateStr = Regex.Match(cellValue.ToString().Trim(), enrollementOrderDatePattern).ToString();

                                            if (DateTime.TryParseExact(
                                                   dateStr,
                                                   "dd.MM.yyyy",
                                                   CultureInfo.InvariantCulture,
                                                   DateTimeStyles.None,
                                                   out DateTime parsedDate))
                                            {
                                                student.EnrollementOrderDate = DateOnly.FromDateTime(parsedDate);
                                            }
                                            else throw new FormatException($"Не удалось распознать дату: {dateStr}");

                                            student.EnrollementOrderNum = Regex.Match(cellValue.ToString(), enrollementOrderNumPattern).ToString();
                                        }
                                        break;
                                }
                            }
                        }
                    }
                }
            }
            return students;
        }

        public async Task<IEnumerable<Student>> GetStudentsEducationPlan(IFormFile plan, List<Student> students)
        {
            using (var stream = new MemoryStream())
            {
                await plan.CopyToAsync(stream);

                using (var package = new ExcelPackage(stream))
                {
                    if (package.Workbook.Worksheets.Count == 0)
                        throw new InvalidOperationException("Загруженный файл не содержит листов");

                    // Извлечение типа контроля и часов из листов Курс 1…4
                    // Столбец 5 — название дисциплины
                    // Столбец 7 — вид контроля для нечётных семестров (1, 3, 5, 7)
                    // Столбец 22 — вид контроля для чётных семестров (2, 4, 6, 8)
                    // Столбец 8 — всего академических часов для нечётных семестров
                    // Столбец 9 — контактные (аудиторные) часы для нечётных семестров
                    // Столбец 23 — всего академических часов для чётных семестров
                    // Столбец 24 — контактные (аудиторные) часы для чётных семестров
                    static string NormalizeControlType(string? value)
                    {
                        if (string.IsNullOrWhiteSpace(value))
                            return string.Empty;

                        var normalized = value.Trim().ToLower();

                        // Эк КР -> экзамен, курсовая работа
                        if (normalized.Contains("эк") && normalized.Contains("кр"))
                            return "экзамен, курсовая";

                        return normalized switch
                        {
                            "эк" => "экзамен",
                            "к" => "контрольная работа",
                            "кр" => "курсовая",
                            "за" => "зачёт",
                            "зао" => "зачёт с оценкой",
                            _ => value.Trim()
                        };
                    }

                    var courseSheets = new[] { "Курс 1", "Курс 2", "Курс 3", "Курс 4" };
                    for (int courseIndex = 0; courseIndex < courseSheets.Length; courseIndex++)
                    {
                        var courseSheet = package.Workbook.Worksheets[courseSheets[courseIndex]];
                        if (courseSheet?.Dimension == null)
                            continue;

                        var courseNum = courseIndex + 1;
                        var oddSemester = (short)(courseNum * 2 - 1);  // 1, 3, 5, 7
                        var evenSemester = (short)(courseNum * 2);     // 2, 4, 6, 8

                        var courseRowCount = courseSheet.Dimension.Rows;

                        // Строки дисциплин на листах «Курс N» начинаются с 6-й (как на ПланСвод)
                        for (int row = 6; row <= courseRowCount; row++)
                        {
                            var nameRaw = courseSheet.Cells[row, 5].Text;
                            var disciplineName = NormalizeDisciplineName(nameRaw);
                            if (string.IsNullOrWhiteSpace(disciplineName))
                                continue;

                            // Вид контроля для нечётного семестра (столбец 7)
                            var oddControlRaw = courseSheet.Cells[row, 7].Text;
                            var oddControlType = NormalizeControlType(oddControlRaw);
                            var hasOddTotalHours = TryGetIntFromCourseSheetCell(courseSheet, row, 8, out var oddTotalHours);
                            var hasOddContactHours = TryGetIntFromCourseSheetCell(courseSheet, row, 9, out var oddContactHours);

                            // Вид контроля для чётного семестра (столбец 22)
                            var evenControlRaw = courseSheet.Cells[row, 22].Text;
                            var evenControlType = NormalizeControlType(evenControlRaw);
                            var hasEvenTotalHours = TryGetIntFromCourseSheetCell(courseSheet, row, 23, out var evenTotalHours);
                            var hasEvenContactHours = TryGetIntFromCourseSheetCell(courseSheet, row, 24, out var evenContactHours);

                            var isBoldRow = courseSheet.Cells[row, 5].Style.Font.Bold;
                            var hasAnyAcademicData = !string.IsNullOrWhiteSpace(oddControlType) ||
                                                     !string.IsNullOrWhiteSpace(evenControlType) ||
                                                     hasOddTotalHours || hasEvenTotalHours ||
                                                     hasOddContactHours || hasEvenContactHours;
                            // Жирные заголовки модулей без данных пропускаем; реальные дисциплины (в т.ч. НИР) — обрабатываем
                            if (isBoldRow && !hasAnyAcademicData)
                                continue;

                            foreach (var student in students)
                            {
                                var targetOdd = student.DisciplineResults.FirstOrDefault(x =>
                                    x.DisciplineName != null &&
                                    x.Semester.HasValue &&
                                    x.Semester.Value == oddSemester &&
                                    NormalizeDisciplineNameForComparison(x.DisciplineName).Equals(
                                        NormalizeDisciplineNameForComparison(disciplineName),
                                        StringComparison.OrdinalIgnoreCase));

                                var targetEven = student.DisciplineResults.FirstOrDefault(x =>
                                    x.DisciplineName != null &&
                                    x.Semester.HasValue &&
                                    x.Semester.Value == evenSemester &&
                                    NormalizeDisciplineNameForComparison(x.DisciplineName).Equals(
                                        NormalizeDisciplineNameForComparison(disciplineName),
                                        StringComparison.OrdinalIgnoreCase));

                                if (targetOdd != null && !string.IsNullOrWhiteSpace(oddControlType) &&
                                    string.IsNullOrWhiteSpace(targetOdd.ControlType))
                                {
                                    targetOdd.ControlType = oddControlType;
                                }

                                if (targetEven != null && !string.IsNullOrWhiteSpace(evenControlType) &&
                                    string.IsNullOrWhiteSpace(targetEven.ControlType))
                                {
                                    targetEven.ControlType = evenControlType;
                                }

                                if (targetOdd != null && hasOddTotalHours && !targetOdd.TotalHours.HasValue)
                                {
                                    targetOdd.TotalHours = oddTotalHours;
                                }

                                if (targetEven != null && hasEvenTotalHours && !targetEven.TotalHours.HasValue)
                                {
                                    targetEven.TotalHours = evenTotalHours;
                                }

                                if (targetOdd != null && hasOddContactHours && !targetOdd.AudHours.HasValue)
                                {
                                    targetOdd.AudHours = oddContactHours;
                                }

                                if (targetEven != null && hasEvenContactHours && !targetEven.AudHours.HasValue)
                                {
                                    targetEven.AudHours = evenContactHours;
                                }
                            }
                        }
                    }
                }
            }

            return students;
        }

        private static string GetMergedCellTextFromWorksheet(OfficeOpenXml.ExcelWorksheet worksheet, int row, int col)
        {
            var mergedAddress = worksheet.MergedCells[row, col];
            if (!string.IsNullOrWhiteSpace(mergedAddress))
                return worksheet.Cells[mergedAddress].First().Text;
            return worksheet.Cells[row, col].Text;
        }

        private static bool TryGetIntFromCourseSheetCell(OfficeOpenXml.ExcelWorksheet worksheet, int row, int col, out int result)
        {
            result = default;
            var text = GetMergedCellTextFromWorksheet(worksheet, row, col)?.Trim();
            if (!string.IsNullOrWhiteSpace(text))
            {
                text = text.Replace(" ", string.Empty).Replace(",", ".");
                if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed))
                {
                    result = (int)Math.Round(parsed);
                    return true;
                }
            }

            var value = worksheet.Cells[row, col].Value;
            if (value == null)
                return false;

            if (value is int i)
            {
                result = i;
                return true;
            }
            if (value is long l)
            {
                result = (int)l;
                return true;
            }
            if (value is double d)
            {
                result = (int)Math.Round(d);
                return true;
            }
            if (value is decimal dec)
            {
                result = (int)Math.Round((double)dec);
                return true;
            }

            var str = value.ToString()?.Trim();
            if (string.IsNullOrWhiteSpace(str))
                return false;
            str = str.Replace(" ", string.Empty).Replace(",", ".");
            if (double.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsedFromValue))
            {
                result = (int)Math.Round(parsedFromValue);
                return true;
            }

            return false;
        }

        /// <summary>Лист семейства «Практическая подготовка» (нейминг может отличаться по пробелам).</summary>
        private static OfficeOpenXml.ExcelWorksheet? FindPracticePreparationWorksheet(ExcelWorkbook workbook)
        {
            foreach (var sheet in workbook.Worksheets)
            {
                var n = Regex.Replace(sheet?.Name ?? "", @"\s+", " ").Trim().ToLowerInvariant();
                if (n.Contains("практическ") && n.Contains("подготовк"))
                    return sheet;
            }

            return null;
        }

        /// <summary>Практики/ФТД с листа «Практическая подготовка» если по ПланСвод не установилась секция.</summary>
        private void OverlayPlanBucketsFromPracticePreparationSheet(ExcelPackage package, List<PlanDisciplineEntry> entries, bool enableFacultyParsing)
        {
            var ws = FindPracticePreparationWorksheet(package.Workbook);
            if (ws?.Dimension == null || entries.Count == 0)
                return;

            var nameToBucket = new Dictionary<string, SupplementPlanBucket>(StringComparer.OrdinalIgnoreCase);
            SupplementPlanBucket currentBucket = SupplementPlanBucket.Discipline;

            int rowEnd = ws.Dimension.Rows;
            for (int row = 4; row <= rowEnd; row++)
            {
                var anchorTxt = NormalizeDisciplineName(GetMergedCellTextFromWorksheet(ws, row, 1));
                if (!string.IsNullOrWhiteSpace(anchorTxt))
                {
                    var sec = ResolvePlanSectionFromAnchor(anchorTxt);
                    if (sec != SupplementPlanBucket.Unknown)
                        currentBucket = sec;
                }

                var nameRaw = ws.Cells[row, 3].Text;
                var disciplineName = NormalizeDisciplineName(nameRaw);
                if (string.IsNullOrWhiteSpace(disciplineName))
                    continue;

                try
                {
                    if (ws.Cells[row, 3].Style.Font.Bold)
                        continue;
                }
                catch { }

                var inPlanMark = ws.Cells[row, 1].Text?.Replace('\u00A0', ' ').Trim() ?? "";
                if (!IsSupplementPlanIncludedByInPlanMarker(inPlanMark))
                    continue;

                var codeCell = SupplementPlanCodeCellFirstLine(ws.Cells[row, 2].Text);
                if (!LooksLikeSupplementPlanDisciplineCode(codeCell))
                    continue;

                if (IsSupplementPlanNoiseName(disciplineName))
                    continue;

                if (currentBucket == SupplementPlanBucket.Elective && !enableFacultyParsing)
                    continue;

                if (currentBucket != SupplementPlanBucket.Practice && currentBucket != SupplementPlanBucket.Elective)
                    continue;

                nameToBucket[disciplineName] = currentBucket;
            }

            foreach (var entry in entries)
            {
                if (entry.PlanBucket != SupplementPlanBucket.Discipline && entry.PlanBucket != SupplementPlanBucket.Unknown)
                    continue;

                var k = NormalizeDisciplineName(entry.DisciplineName);
                if (string.IsNullOrWhiteSpace(k))
                    continue;

                if (!nameToBucket.TryGetValue(k, out var overlayBucket))
                    continue;

                if (overlayBucket == SupplementPlanBucket.Elective && !enableFacultyParsing)
                    continue;

                entry.PlanBucket = overlayBucket;
            }
        }

        private static SupplementPlanBucket ResolvePlanSectionFromAnchor(string? raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
                return SupplementPlanBucket.Unknown;
            var t = Regex.Replace(raw, @"\s+", " ").Trim().ToLowerInvariant();
            if (t.Contains("фтд") && (t.Contains("факультат") || t.Contains("факультатив")))
                return SupplementPlanBucket.Elective;
            var compact = Regex.Replace(t, @"\s+", "");
            if (Regex.IsMatch(compact, @"блок3[\.\)].*(аттест|государств|итогов)",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                return SupplementPlanBucket.Gia;
            if (t.Contains("блок") && t.Contains("2") && (t.Contains("практик") || t.Contains("практика")))
                return SupplementPlanBucket.Practice;
            if (t.Contains("блок") && t.Contains("1"))
                return SupplementPlanBucket.Discipline;
            return SupplementPlanBucket.Unknown;
        }

        private static bool LooksLikeSupplementPlanDisciplineCode(string? code)
        {
            code = SupplementPlanCodeCellFirstLine(code);
            if (string.IsNullOrWhiteSpace(code))
                return false;
            code = code.Trim();
            if (code.Length == 0)
                return false;
            var c0 = char.ToUpperInvariant(code[0]);
            if (c0 == '\u0411')
                return true;
            if (c0 == 'B')
                return true;
            if (code.StartsWith("ФТД", StringComparison.OrdinalIgnoreCase))
                return true;
            return false;
        }

        private static bool IsSupplementPlanNoiseName(string normalizedName)
        {
            if (string.IsNullOrWhiteSpace(normalizedName))
                return true;
            var lower = normalizedName.ToLowerInvariant();
            if (lower.Contains("итого") && lower.Contains("семестр"))
                return true;
            if (lower.StartsWith("итого", StringComparison.OrdinalIgnoreCase))
                return true;
            if (string.Equals(lower, "в том числе:", StringComparison.OrdinalIgnoreCase))
                return true;
            if (string.Equals(lower, "в том числе", StringComparison.OrdinalIgnoreCase))
                return true;
            return false;
        }

        /// <summary>Строки из карточки, которые не должны попадать в итог (шум без соответствия плану).</summary>
        private static bool SkipLeftoverDisciplineAfterPlan(string? rawName)
        {
            var norm = NormalizeDisciplineName(rawName ?? "");
            if (string.IsNullOrWhiteSpace(norm))
                return true;
            if (IsSupplementPlanNoiseName(norm))
                return true;
            return false;
        }

        /// <summary>Ячейка "Свод" L8 - целевой объём ОП в з.е.</summary>
        private static async Task<double?> ReadTargetProgramCreditsFromPlanAsync(IFormFile plan)
        {
            using var stream = new MemoryStream();
            await plan.CopyToAsync(stream);
            stream.Position = 0;
            using var package = new ExcelPackage(stream);
            var ws = package.Workbook.Worksheets["Свод"];
            if (ws == null)
                return null;

            var value = ws.Cells[8, 12].Value;
            if (value == null)
                return null;
            if (value is double d)
                return d;
            if (value is float f)
                return f;
            if (value is decimal dec)
                return (double)dec;
            if (value is int i)
                return i;
            if (value is long l)
                return l;

            var str = value.ToString()?.Trim();
            if (string.IsNullOrWhiteSpace(str))
                return null;
            str = str.Replace(" ", string.Empty).Replace(",", ".");
            if (double.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed))
                return parsed;
            return null;
        }

        private static bool IsPlanSvodBlock1TotalsAnchorLower(string anchorLower)
        {
            if (string.IsNullOrWhiteSpace(anchorLower))
                return false;
            var compact = Regex.Replace(anchorLower, @"\s+", "", RegexOptions.CultureInvariant);
            return Regex.IsMatch(compact, @"блок1[\.\)].*(диск|модул)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }

        private static bool IsPlanSvodBlock2TotalsAnchorLower(string anchorLower)
        {
            if (string.IsNullOrWhiteSpace(anchorLower))
                return false;
            var compact = Regex.Replace(anchorLower, @"\s+", "", RegexOptions.CultureInvariant);
            return Regex.IsMatch(compact, @"блок2[\.\)].*практ", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }

        private static bool IsPlanSvodBlock3TotalsAnchorLower(string anchorLower)
        {
            if (string.IsNullOrWhiteSpace(anchorLower))
                return false;
            var compact = Regex.Replace(anchorLower, @"\s+", "", RegexOptions.CultureInvariant);
            return Regex.IsMatch(compact, @"блок3[\.\)].*(аттест|государств|итогов)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }

        /// <summary>Сумма строк «Конт. раб.» под заголовками блоков 1–3 листа «ПланСвод».</summary>
        private static async Task<double?> ReadTargetContactHoursFromPlanBlockSummariesAsync(IFormFile plan)
        {
            using var stream = new MemoryStream();
            await plan.CopyToAsync(stream);
            stream.Position = 0;

            using var package = new ExcelPackage(stream);
            try
            {
                package.Workbook.Calculate();
            }
            catch
            {
                // как в карточке: сломанные ссылки — остаётся Text
            }

            var worksheet = package.Workbook.Worksheets["ПланСвод"];
            if (worksheet?.Dimension == null)
                return null;

            static string NormalizeHeaderCell(string? value)
            {
                if (string.IsNullOrWhiteSpace(value))
                    return string.Empty;
                return Regex.Replace(value, @"\s+", " ").Trim();
            }

            static string NormalizeHeaderSquash(string value)
                => NormalizeHeaderCell(value).ToLowerInvariant().Replace(" ", "");

            string KontHeaderText(int row, int col)
            {
                var mergedAddress = worksheet.MergedCells[row, col];
                if (!string.IsNullOrWhiteSpace(mergedAddress))
                    return worksheet.Cells[mergedAddress].First().Text;
                return worksheet.Cells[row, col].Text;
            }

            bool TryNumericFromCell(int row, int col, out double r)
            {
                r = default;
                var cell = worksheet.Cells[row, col];
                var value = cell.Value;

                double FromObj(object v)
                {
                    if (v == null)
                        return double.NaN;
                    if (v is double d)
                        return d;
                    if (v is float f)
                        return f;
                    if (v is decimal dec)
                        return (double)dec;
                    if (v is int i)
                        return i;
                    if (v is long l)
                        return l;
                    var str = v.ToString()?.Trim();
                    if (string.IsNullOrWhiteSpace(str))
                        return double.NaN;
                    str = str.Replace(" ", string.Empty).Replace(",", ".");
                    return double.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out var pv) ? pv : double.NaN;
                }

                var pv = FromObj(value);
                if (!double.IsNaN(pv))
                {
                    r = pv;
                    return true;
                }

                var tx = cell.Text?.Trim().Replace('\u00A0', ' ');
                if (string.IsNullOrWhiteSpace(tx))
                    return false;
                tx = tx.Replace(" ", string.Empty).Replace(",", ".");
                return double.TryParse(tx, NumberStyles.Any, CultureInfo.InvariantCulture, out r);
            }

            string GetMergedText(int row, int col)
            {
                var mergedAddress = worksheet.MergedCells[row, col];
                if (!string.IsNullOrWhiteSpace(mergedAddress))
                    return worksheet.Cells[mergedAddress].First().Text;
                return worksheet.Cells[row, col].Text;
            }

            int columnCount = worksheet.Dimension.Columns;
            int rowCount = worksheet.Dimension.Rows;
            int? kontRabCol = null;

            for (int col = 1; col <= columnCount; col++)
            {
                var header1 = NormalizeHeaderSquash(GetMergedText(1, col));
                var header2 = NormalizeHeaderSquash(GetMergedText(2, col));
                var mergedTopNotes = $"{header1} {header2}".Trim();

                var h3Raw = NormalizeHeaderCell(KontHeaderText(3, col));
                var h3 = NormalizeHeaderSquash(h3Raw);

                bool topLooksItogo =
                    mergedTopNotes.Contains("итого", StringComparison.OrdinalIgnoreCase)
                    || mergedTopNotes.Contains("всего", StringComparison.OrdinalIgnoreCase);

                bool looksKont =
                    Regex.IsMatch(h3, @"конт\.раб", RegexOptions.IgnoreCase)
                    || (h3.Contains("конт") && h3.Contains("раб") && !h3.Contains("роль"))
                    ;

                if (topLooksItogo && looksKont)
                {
                    kontRabCol = col;
                    break;
                }
            }

            if (kontRabCol == null)
            {
                for (int col = 1; col <= columnCount; col++)
                {
                    var h3 = NormalizeHeaderSquash(KontHeaderText(3, col));
                    bool looksKont =
                        Regex.IsMatch(h3, @"конт\.раб", RegexOptions.IgnoreCase)
                        || (h3.Contains("конт") && h3.Contains("раб") && !h3.Contains("роль"))
                        ;
                    if (looksKont)
                    {
                        kontRabCol = col;
                        break;
                    }
                }
            }

            if (kontRabCol == null)
                return null;

            bool got1 = false, got2 = false, got3 = false;
            double sum = 0d;

            for (int row = 4; row <= rowCount; row++)
            {
                string anchorRaw;
                var mergedAddrA = worksheet.MergedCells[row, 1];
                if (!string.IsNullOrWhiteSpace(mergedAddrA))
                    anchorRaw = worksheet.Cells[mergedAddrA].First().Text;
                else
                    anchorRaw = worksheet.Cells[row, 1].Text;

                var anchorLower = NormalizeDisciplineName(anchorRaw).ToLowerInvariant();

                if (!IsPlanSvodBlock1TotalsAnchorLower(anchorLower)
                    && !IsPlanSvodBlock2TotalsAnchorLower(anchorLower)
                    && !IsPlanSvodBlock3TotalsAnchorLower(anchorLower))
                {
                    continue;
                }

                if (!TryNumericFromCell(row, kontRabCol.Value, out var cellVal))
                    continue;

                if (IsPlanSvodBlock1TotalsAnchorLower(anchorLower) && !got1)
                {
                    sum += cellVal;
                    got1 = true;
                }
                else if (IsPlanSvodBlock2TotalsAnchorLower(anchorLower) && !got2)
                {
                    sum += cellVal;
                    got2 = true;
                }
                else if (IsPlanSvodBlock3TotalsAnchorLower(anchorLower) && !got3)
                {
                    sum += cellVal;
                    got3 = true;
                }

                if (got1 && got2 && got3)
                    break;
            }

            if (!got1 || !got2 || !got3)
                return null;
            return sum;
        }

        /// <summary>Итого з.е. блока ГИА (ячейка «Факт» в строке-заголовке «Блок 3 …»).</summary>
        private static async Task<double?> ReadTargetGiaCreditsFromPlanAsync(IFormFile plan)
        {
            using var stream = new MemoryStream();
            await plan.CopyToAsync(stream);
            stream.Position = 0;

            using var package = new ExcelPackage(stream);
            try
            {
                package.Workbook.Calculate();
            }
            catch
            {
            }

            var worksheet = package.Workbook.Worksheets["ПланСвод"];
            if (worksheet?.Dimension == null)
                return null;

            static string NormalizeHeaderCell(string? value)
            {
                if (string.IsNullOrWhiteSpace(value))
                    return string.Empty;
                return Regex.Replace(value, @"\s+", " ").Trim();
            }

            static string NormalizeHeaderSquash(string value)
                => NormalizeHeaderCell(value).ToLowerInvariant().Replace(" ", string.Empty);

            string KontHeaderText(int row, int col)
            {
                var mergedAddress = worksheet.MergedCells[row, col];
                if (!string.IsNullOrWhiteSpace(mergedAddress))
                    return worksheet.Cells[mergedAddress].First().Text;
                return worksheet.Cells[row, col].Text;
            }

            bool TryNumericFromCell(int row, int col, out double r)
            {
                r = default;
                var cell = worksheet.Cells[row, col];
                var value = cell.Value;

                double FromObj(object v)
                {
                    if (v == null)
                        return double.NaN;
                    if (v is double d)
                        return d;
                    if (v is float f)
                        return f;
                    if (v is decimal dec)
                        return (double)dec;
                    if (v is int i)
                        return i;
                    if (v is long l)
                        return l;
                    var str = v.ToString()?.Trim();
                    if (string.IsNullOrWhiteSpace(str))
                        return double.NaN;
                    str = str.Replace(" ", string.Empty).Replace(",", ".");
                    return double.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out var pv) ? pv : double.NaN;
                }

                var pv = FromObj(value);
                if (!double.IsNaN(pv))
                {
                    r = pv;
                    return true;
                }

                var tx = cell.Text?.Trim().Replace('\u00A0', ' ');
                if (string.IsNullOrWhiteSpace(tx))
                    return false;
                tx = tx.Replace(" ", string.Empty).Replace(",", ".");
                return double.TryParse(tx, NumberStyles.Any, CultureInfo.InvariantCulture, out r);
            }

            int columnCount = worksheet.Dimension.Columns;
            int rowCount = worksheet.Dimension.Rows;
            int? faktZeCol = null;

            for (int col = 1; col <= columnCount; col++)
            {
                var h3 = NormalizeHeaderSquash(KontHeaderText(3, col));
                if (h3.Length == 0)
                    continue;
                // «Факт» под з.е.; не брать заголовки вроде «факультатив»
                if (h3 == "факт")
                {
                    faktZeCol = col;
                    break;
                }
            }

            if (faktZeCol == null)
                return null;

            for (int row = 4; row <= rowCount; row++)
            {
                string anchorRaw;
                var mergedAddrA = worksheet.MergedCells[row, 1];
                if (!string.IsNullOrWhiteSpace(mergedAddrA))
                    anchorRaw = worksheet.Cells[mergedAddrA].First().Text;
                else
                    anchorRaw = worksheet.Cells[row, 1].Text;

                var anchorLower = NormalizeDisciplineName(anchorRaw).ToLowerInvariant();
                if (!IsPlanSvodBlock3TotalsAnchorLower(anchorLower))
                    continue;

                if (TryNumericFromCell(row, faktZeCol.Value, out var v))
                    return v;
                return null;
            }

            return null;
        }

        /// <summary>Итого з.е. блока «Практика» из строки-заголовка «Блок 2 …» (колонка «Факт»).</summary>
        private static async Task<double?> ReadTargetPracticeCreditsFromPlanAsync(IFormFile plan)
        {
            using var stream = new MemoryStream();
            await plan.CopyToAsync(stream);
            stream.Position = 0;

            using var package = new ExcelPackage(stream);
            try
            {
                package.Workbook.Calculate();
            }
            catch
            {
            }

            var worksheet = package.Workbook.Worksheets["ПланСвод"];
            if (worksheet?.Dimension == null)
                return null;

            static string NormalizeHeaderCell(string? value)
            {
                if (string.IsNullOrWhiteSpace(value))
                    return string.Empty;
                return Regex.Replace(value, @"\s+", " ").Trim();
            }

            static string NormalizeHeaderSquash(string value)
                => NormalizeHeaderCell(value).ToLowerInvariant().Replace(" ", string.Empty);

            string KontHeaderText(int row, int col)
            {
                var mergedAddress = worksheet.MergedCells[row, col];
                if (!string.IsNullOrWhiteSpace(mergedAddress))
                    return worksheet.Cells[mergedAddress].First().Text;
                return worksheet.Cells[row, col].Text;
            }

            bool TryNumericFromCell(int row, int col, out double r)
            {
                r = default;
                var cell = worksheet.Cells[row, col];
                var value = cell.Value;

                double FromObj(object v)
                {
                    if (v == null)
                        return double.NaN;
                    if (v is double d)
                        return d;
                    if (v is float f)
                        return f;
                    if (v is decimal dec)
                        return (double)dec;
                    if (v is int i)
                        return i;
                    if (v is long l)
                        return l;
                    var str = v.ToString()?.Trim();
                    if (string.IsNullOrWhiteSpace(str))
                        return double.NaN;
                    str = str.Replace(" ", string.Empty).Replace(",", ".");
                    return double.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out var pv) ? pv : double.NaN;
                }

                var pv = FromObj(value);
                if (!double.IsNaN(pv))
                {
                    r = pv;
                    return true;
                }

                var tx = cell.Text?.Trim().Replace('\u00A0', ' ');
                if (string.IsNullOrWhiteSpace(tx))
                    return false;
                tx = tx.Replace(" ", string.Empty).Replace(",", ".");
                return double.TryParse(tx, NumberStyles.Any, CultureInfo.InvariantCulture, out r);
            }

            int columnCount = worksheet.Dimension.Columns;
            int rowCount = worksheet.Dimension.Rows;
            int? faktZeCol = null;

            for (int col = 1; col <= columnCount; col++)
            {
                var h3 = NormalizeHeaderSquash(KontHeaderText(3, col));
                if (h3.Length == 0)
                    continue;
                if (h3 == "факт")
                {
                    faktZeCol = col;
                    break;
                }
            }

            if (faktZeCol == null)
                return null;

            for (int row = 4; row <= rowCount; row++)
            {
                string anchorRaw;
                var mergedAddrA = worksheet.MergedCells[row, 1];
                if (!string.IsNullOrWhiteSpace(mergedAddrA))
                    anchorRaw = worksheet.Cells[mergedAddrA].First().Text;
                else
                    anchorRaw = worksheet.Cells[row, 1].Text;

                var anchorLower = NormalizeDisciplineName(anchorRaw).ToLowerInvariant();
                if (!IsPlanSvodBlock2TotalsAnchorLower(anchorLower))
                    continue;

                if (TryNumericFromCell(row, faktZeCol.Value, out var v))
                    return v;
                return null;
            }

            return null;
        }

        private async Task<List<PlanDisciplineEntry>> ParseStudyPlanForSupplementAsync(IFormFile plan, bool enableFacultyParsing)
        {
            var result = new List<PlanDisciplineEntry>();
            using (var stream = new MemoryStream())
            {
                await plan.CopyToAsync(stream);

                using (var package = new ExcelPackage(stream))
                {
                    if (package.Workbook.Worksheets.Count == 0)
                        throw new InvalidOperationException("Загруженный файл не содержит листов");

                    var worksheet = package.Workbook.Worksheets["ПланСвод"]
                        ?? throw new InvalidOperationException("Загруженный файл не содержит листа ПланСвод");

                    if (worksheet.Dimension == null)
                    {
                        logger.LogWarning("ПланСвод: Dimension=null, файл={FileName}, пустой результат.", plan.FileName);
                        return result;
                    }

                    var columnCount = worksheet.Dimension.Columns;
                    var rowCount = worksheet.Dimension.Rows;

                    static string PlanColLetter(int columnNumber)
                    {
                        if (columnNumber < 1)
                            return "?";
                        string letters = "";
                        var n = columnNumber;
                        while (n > 0)
                        {
                            n--;
                            letters = (char)('A' + n % 26) + letters;
                            n /= 26;
                        }
                        return letters;
                    }

                    logger.LogInformation(
                        "ПланСвод: старт разбора листа «ПланСвод», файл={FileName}, строк данных до {RowMax}, колонок {ColMax}, EnableFacultyParsing={Faculty}",
                        plan.FileName, rowCount, columnCount, enableFacultyParsing);

                    var currentSectionBucket = SupplementPlanBucket.Discipline;
                    string? practiceTypePrefixBlock2 = null;

                    string GetMergedText(int row, int col)
                    {
                        var mergedAddress = worksheet.MergedCells[row, col];
                        if (!string.IsNullOrWhiteSpace(mergedAddress))
                            return worksheet.Cells[mergedAddress].First().Text;
                        return worksheet.Cells[row, col].Text;
                    }

                    static string LocalNormalizePlanHeader(string? value)
                    {
                        if (string.IsNullOrWhiteSpace(value))
                            return string.Empty;
                        return Regex.Replace(value, @"\s+", " ").Trim();
                    }

                    int? audHoursCol = null;
                    var creditUnitsBySemesterCol = new Dictionary<int, int>();
                    int? inPlanCol = null;
                    int? codeCol = null;
                    int? nameCol = null;

                    for (int col = 1; col <= columnCount; col++)
                    {
                        var header2 = LocalNormalizePlanHeader(GetMergedText(2, col));
                        var header3 = LocalNormalizePlanHeader(GetMergedText(3, col));

                        var header2Lower = header2.ToLower();
                        var header3Lower = header3.ToLower();

                        if (inPlanCol == null && (header3Lower.Contains("считать") || header3Lower.Contains("в плане")))
                        {
                            inPlanCol = col;
                            logger.LogDebug(
                                "ПланСвод: колонка «Считать в плане» → {ColLetter} ({Col}), причина: строка3 содержит «считать»/«в плане», текст=\"{Header3}\"",
                                PlanColLetter(col), col, header3);
                        }
                        else if (codeCol == null && header3Lower.Contains("индекс"))
                        {
                            codeCol = col;
                            logger.LogDebug(
                                "ПланСвод: колонка «Индекс» → {ColLetter} ({Col}), причина: строка3 содержит «индекс», текст=\"{Header3}\"",
                                PlanColLetter(col), col, header3);
                        }
                        else if (nameCol == null && header3Lower.Contains("наименование"))
                        {
                            nameCol = col;
                            logger.LogDebug(
                                "ПланСвод: колонка «Наименование» → {ColLetter} ({Col}), причина: строка3 содержит «наименование», текст=\"{Header3}\"",
                                PlanColLetter(col), col, header3);
                        }

                        if (audHoursCol == null &&
                            (header2Lower.Contains("итого") || header2Lower.Contains("всего")) &&
                            Regex.IsMatch(header3Lower, @"\bауд\.?\b", RegexOptions.IgnoreCase))
                        {
                            audHoursCol = col;
                            logger.LogDebug(
                                "ПланСвод: колонка итого ауд. часов → {ColLetter} ({Col}): строка2=\"{H2}\", строка3=\"{H3}\"",
                                PlanColLetter(col), col, header2, header3);
                            continue;
                        }

                        var matchSemester = Regex.Match(header2Lower, @"семестр\s*(\d{1,2})");
                        if (matchSemester.Success && (header3Lower.Contains("з.е") || header3Lower.Contains("з. е")))
                        {
                            if (int.TryParse(matchSemester.Groups[1].Value, out var sem) && sem >= 1 && sem <= 12)
                            {
                                creditUnitsBySemesterCol[sem] = col;
                                logger.LogDebug(
                                    "ПланСвод: з.е. семестр {Sem} → колонка {ColLetter} ({Col}), строка2=\"{H2}\", строка3=\"{H3}\"",
                                    sem, PlanColLetter(col), col, header2, header3);
                            }
                        }
                        else if (Regex.IsMatch(header2Lower, @"семестр\s*[аАaA]\b") &&
                            (header3Lower.Contains("з.е") || header3Lower.Contains("з. е")))
                        {
                            creditUnitsBySemesterCol[10] = col;
                            logger.LogDebug(
                                "ПланСвод: з.е. семестр 10 (буква А) → колонка {ColLetter} ({Col}), строка2=\"{H2}\"",
                                PlanColLetter(col), col, header2);
                        }
                    }

                    if (inPlanCol == null || codeCol == null || nameCol == null)
                        throw new InvalidOperationException("Не удалось определить колонки ПланСвод (Считать в плане / Индекс / Наименование)");

                    if (audHoursCol == null || creditUnitsBySemesterCol.Count == 0)
                        throw new InvalidOperationException("Не удалось определить колонки учебного плана (Ауд. часов/з.е. по семестрам)");

                    var semMapForLog = string.Join(", ",
                        creditUnitsBySemesterCol.OrderBy(kv => kv.Key).Select(kv =>
                            $"семестр{kv.Key}={PlanColLetter(kv.Value)}{kv.Value}"));
                    logger.LogInformation(
                        "ПланСвод: карта колонок принята — СчитатьВПлане={InPlan} ({InPlanL}), Индекс={Code} ({CodeL}), Наименование={Name} ({NameL}), АудИтого={Aud} ({AudL}); з.е.: {SemMap}",
                        inPlanCol, PlanColLetter(inPlanCol.Value), codeCol, PlanColLetter(codeCol.Value),
                        nameCol, PlanColLetter(nameCol.Value), audHoursCol, PlanColLetter(audHoursCol.Value), semMapForLog);

                    bool TryGetIntCell(int row, int col, out int r)
                    {
                        r = default;
                        var value = worksheet.Cells[row, col].Value;
                        if (value == null)
                            return false;

                        if (value is int i)
                        {
                            r = i;
                            return true;
                        }
                        if (value is long l)
                        {
                            r = (int)l;
                            return true;
                        }
                        if (value is double d)
                        {
                            r = (int)Math.Round(d);
                            return true;
                        }
                        if (value is decimal dec)
                        {
                            r = (int)Math.Round((double)dec);
                            return true;
                        }

                        var str = value.ToString()?.Trim();
                        if (string.IsNullOrWhiteSpace(str))
                            return false;
                        str = str.Replace(" ", string.Empty).Replace(",", ".");
                        if (double.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed))
                        {
                            r = (int)Math.Round(parsed);
                            return true;
                        }
                        return false;
                    }

                    bool TryGetDoubleCell(int row, int col, out double r)
                    {
                        r = default;
                        var value = worksheet.Cells[row, col].Value;
                        if (value == null)
                            return false;

                        if (value is double d)
                        {
                            r = d;
                            return true;
                        }
                        if (value is float f)
                        {
                            r = f;
                            return true;
                        }
                        if (value is decimal dec)
                        {
                            r = (double)dec;
                            return true;
                        }
                        if (value is int i)
                        {
                            r = i;
                            return true;
                        }
                        if (value is long l)
                        {
                            r = l;
                            return true;
                        }

                        var str = value.ToString()?.Trim();
                        if (string.IsNullOrWhiteSpace(str))
                            return false;
                        str = str.Replace(" ", string.Empty).Replace(",", ".");
                        return double.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out r);
                    }

                    var nSkipEmptyName = 0;
                    var nSkipPracticeBoldPrefix = 0;
                    var nSkipBoldHeader = 0;
                    var nSkipNoPlus = 0;
                    var nSkipBadCode = 0;
                    var nSkipNoise = 0;
                    var nSkipGiaExcluded = 0;
                    var nSkipElectiveFtd = 0;

                    for (int row = 6; row <= rowCount; row++)
                    {
                        var anchorTxt = NormalizeDisciplineName(GetMergedText(row, 1));
                        if (!string.IsNullOrWhiteSpace(anchorTxt))
                        {
                            var sec = ResolvePlanSectionFromAnchor(anchorTxt);
                            if (sec != SupplementPlanBucket.Unknown)
                            {
                                var bucketOld = currentSectionBucket;
                                currentSectionBucket = sec;
                                if (sec != SupplementPlanBucket.Practice)
                                    practiceTypePrefixBlock2 = null;
                                if (bucketOld != currentSectionBucket)
                                {
                                    logger.LogDebug(
                                        "ПланСвод: строка {Row}, колонка A якорь «{Anchor}» → секция {Bucket} (было {Was})",
                                        row, anchorTxt, currentSectionBucket, bucketOld);
                                }
                            }
                        }

                        var nameRaw = worksheet.Cells[row, nameCol.Value].Text;
                        var disciplineName = NormalizeDisciplineName(nameRaw);
                        if (string.IsNullOrWhiteSpace(disciplineName))
                        {
                            nSkipEmptyName++;
                            continue;
                        }

                        bool dBold = false;
                        try { dBold = worksheet.Cells[row, nameCol.Value].Style.Font.Bold; } catch { }

                        if (currentSectionBucket == SupplementPlanBucket.Practice && dBold)
                        {
                            practiceTypePrefixBlock2 = disciplineName;
                            nSkipPracticeBoldPrefix++;
                            logger.LogDebug(
                                "ПланСвод: строка {Row}, зафиксирован префикс практики (жирная строка блока 2): «{Prefix}»",
                                row, disciplineName);
                            continue;
                        }

                        if (dBold)
                        {
                            nSkipBoldHeader++;
                            logger.LogDebug(
                                "ПланСвод: строка {Row}, пропуск жирной строки (заголовок/группа): «{Name}»",
                                row, disciplineName);
                            continue;
                        }

                        var inPlanMark = worksheet.Cells[row, inPlanCol.Value].Text?.Replace('\u00A0', ' ').Trim() ?? string.Empty;
                        if (!IsSupplementPlanIncludedByInPlanMarker(inPlanMark))
                        {
                            nSkipNoPlus++;
                            logger.LogDebug(
                                "ПланСвод: строка {Row}, отметка «в плане» не проходит правило в {InPlanCol} ({InPlanL}), значение=\"{Mark}\", название «{Name}»",
                                row, inPlanCol.Value, PlanColLetter(inPlanCol.Value), inPlanMark, disciplineName);
                            continue;
                        }

                        var codeCell = SupplementPlanCodeCellFirstLine(worksheet.Cells[row, codeCol.Value].Text);
                        if (!LooksLikeSupplementPlanDisciplineCode(codeCell))
                        {
                            nSkipBadCode++;
                            logger.LogDebug(
                                "ПланСвод: строка {Row}, код в {CodeCol} ({CodeL}) не похож на индекс плана: «{Code}», название «{Name}»",
                                row, codeCol.Value, PlanColLetter(codeCol.Value), codeCell, disciplineName);
                            continue;
                        }

                        if (IsSupplementPlanNoiseName(disciplineName))
                        {
                            nSkipNoise++;
                            logger.LogDebug(
                                "ПланСвод: строка {Row}, служебная/итоговая строка по имени: «{Name}»",
                                row, disciplineName);
                            continue;
                        }

                        if (currentSectionBucket == SupplementPlanBucket.Gia &&
                            LooksLikePlanGiaDefensePreparationExcluded(disciplineName))
                        {
                            nSkipGiaExcluded++;
                            logger.LogDebug(
                                "ПланСвод: строка {Row}, ГИА: исключение по правилу подготовки к защите: «{Name}»",
                                row, disciplineName);
                            continue;
                        }

                        double? aud = null;
                        if (TryGetIntCell(row, audHoursCol.Value, out var hours))
                            aud = hours;

                        var bySem = new Dictionary<int, double>();
                        foreach (var kvp in creditUnitsBySemesterCol.OrderBy(x => x.Key))
                        {
                            if (!TryGetDoubleCell(row, kvp.Value, out var ze))
                                continue;
                            bySem[kvp.Key] = ze;
                        }

                        var bucket = currentSectionBucket;

                        if (!enableFacultyParsing &&
                            (bucket == SupplementPlanBucket.Elective ||
                             codeCell.StartsWith("ФТД", StringComparison.OrdinalIgnoreCase)))
                        {
                            nSkipElectiveFtd++;
                            logger.LogDebug(
                                "ПланСвод: строка {Row}, пропуск факультатива/ФТД при EnableFacultyParsing=false: код «{Code}», bucket={Bucket}, «{Name}»",
                                row, codeCell, bucket, disciplineName);
                            continue;
                        }

                        var finalDisciplineName = disciplineName;
                        if (currentSectionBucket == SupplementPlanBucket.Practice &&
                            !string.IsNullOrWhiteSpace(practiceTypePrefixBlock2))
                        {
                            finalDisciplineName = FormatPracticePlanNameWithTypePrefix(
                                practiceTypePrefixBlock2,
                                disciplineName);
                        }

                        var zeSummary = bySem.Count == 0
                            ? "(нет чисел з.е.)"
                            : string.Join(", ", bySem.OrderBy(kv => kv.Key).Select(kv => $"{kv.Key}:{kv.Value}"));

                        logger.LogDebug(
                            "ПланСвод: строка {Row} принята — код «{Code}», имя «{FinalName}», секция={Bucket}, ауд={Aud}, з.е.: {Ze}",
                            row, codeCell, finalDisciplineName, bucket, aud, zeSummary);

                        result.Add(new PlanDisciplineEntry
                        {
                            DisciplineName = finalDisciplineName,
                            PlanOrder = row,
                            TotalAudHours = aud,
                            CreditUnitsBySemester = bySem,
                            PlanBucket = bucket
                        });
                    }

                    logger.LogInformation(
                        "ПланСвод: конец разбора строк {RowFrom}-{RowTo}: записей в план {Added}; пропуски: пустоеНаименование={E}, префиксПрактикиЖирный={Pr}, жирныйЗаголовок={Bold}, нетПлюса={NoPlus}, кодНеИндекс={Bad}, шумИтоги={Noise}, гиаИсключено={Gia}, фтдПриОтклФакультатива={Ftd}",
                        6, rowCount, result.Count, nSkipEmptyName, nSkipPracticeBoldPrefix, nSkipBoldHeader,
                        nSkipNoPlus, nSkipBadCode, nSkipNoise, nSkipGiaExcluded, nSkipElectiveFtd);

                    var approxDataRows = Math.Max(1, rowCount - 5);
                    if (nSkipNoPlus > Math.Max(10, approxDataRows / 3))
                    {
                        logger.LogWarning(
                            "ПланСвод: много строк отфильтровано по колонке «в плане» (счётчик={NoPlus}, ~строк данных={Approx}). Проверьте отметки в файле или Import:SupplementStudyPlan:RequireExplicitPlusInPlan / TreatEmptyInPlanMarkerAsIncluded.",
                            nSkipNoPlus, approxDataRows);
                    }
                    if (nSkipBadCode > Math.Max(10, approxDataRows / 3))
                    {
                        logger.LogWarning(
                            "ПланСвод: много строк с нераспознанным индексом (счётчик={Bad}, ~строк данных={Approx}). Часто в ячейке несколько строк или латинская «B» вместо «Б».",
                            nSkipBadCode, approxDataRows);
                    }
                }
            }

            return result;
        }

        private List<StudentDisciplineResult> BuildDisciplineResultsFromPlanAndCard(
            List<PlanDisciplineEntry> planEntries,
            List<StudentDisciplineResult> cardOnly)
        {
            var pool = new List<StudentDisciplineResult>(cardOnly);
            var output = new List<StudentDisciplineResult>();

            var statExact = 0;
            var statExactAmbiguous = 0;
            var statRelaxed = 0;
            var statRelaxedAmbiguous = 0;
            var statTruncated = 0;
            var statTruncatedAmbiguous = 0;
            var statLoosenedPractice = 0;
            var statLoosenedPracticeAmbiguous = 0;
            var statCourseWorkBySemester = 0;
            var statFuzzy = 0;
            var statFuzzyAmbiguous = 0;
            var statPlanRowNoCard = 0;

            logger.LogInformation(
                "План↔карточка: начало сопоставления, строк в учебном плане={PlanCount}, оценок из карточки (пул)={CardCount}",
                planEntries.Count, cardOnly.Count);

            foreach (var pe in planEntries.OrderBy(p => p.PlanOrder))
            {
                var planName = pe.DisciplineName;
                var totalZe = pe.CreditUnitsBySemester.Values.Sum();
                double? totalCredits = totalZe > 0 ? totalZe : (double?)null;

                short? planGuessSem = null;
                if (pe.CreditUnitsBySemester.Any(kv => kv.Value > 0))
                    planGuessSem = (short)pe.CreditUnitsBySemester.Where(kv => kv.Value > 0).OrderBy(kv => kv.Key).First().Key;

                var exact = pool
                    .Where(c => c.DisciplineName != null && ExactPlanCardNameMatch(planName, c.DisciplineName))
                    .ToList();

                if (exact.Count == 1)
                {
                    statExact++;
                    var chosen = exact[0];
                    pool.Remove(chosen);
                    logger.LogDebug(
                        "План↔карточка: PlanOrder={Order} «{Name}» — совпадение exact с карточкой «{CardName}», оценка={Score}",
                        pe.PlanOrder, planName, chosen.DisciplineName, chosen.Score);
                    output.Add(new StudentDisciplineResult
                    {
                        PlanBucket = pe.PlanBucket,
                        DisciplineName = planName,
                        PlanOrder = pe.PlanOrder,
                        AudHours = pe.TotalAudHours,
                        CreditUnits = totalCredits,
                        Score = chosen.Score,
                        Semester = chosen.Semester ?? planGuessSem,
                        Year = chosen.Year,
                        ControlType = chosen.ControlType,
                        RequiresManualValidation = false
                    });
                    continue;
                }

                if (exact.Count > 1)
                {
                    statExactAmbiguous++;
                    foreach (var t in exact)
                        pool.Remove(t);
                    logger.LogWarning(
                        "План: PlanOrder={Order} «{Name}» — несколько точных совпадений в карточке, оценка не проставляется.",
                        pe.PlanOrder, planName);
                    output.Add(new StudentDisciplineResult
                    {
                        PlanBucket = pe.PlanBucket,
                        DisciplineName = planName,
                        PlanOrder = pe.PlanOrder,
                        AudHours = pe.TotalAudHours,
                        CreditUnits = totalCredits,
                        Semester = planGuessSem,
                        RequiresManualValidation = true
                    });
                    continue;
                }

                var relaxed = pool
                    .Where(c => c.DisciplineName != null && RelaxedPlanCardNameMatch(planName, c.DisciplineName))
                    .ToList();

                if (relaxed.Count == 1)
                {
                    statRelaxed++;
                    var chosen = relaxed[0];
                    pool.Remove(chosen);
                    logger.LogDebug(
                        "План↔карточка: PlanOrder={Order} «{Name}» — совпадение relaxed (нормализация для сравнения) с «{CardName}»",
                        pe.PlanOrder, planName, chosen.DisciplineName);
                    output.Add(new StudentDisciplineResult
                    {
                        PlanBucket = pe.PlanBucket,
                        DisciplineName = planName,
                        PlanOrder = pe.PlanOrder,
                        AudHours = pe.TotalAudHours,
                        CreditUnits = totalCredits,
                        Score = chosen.Score,
                        Semester = chosen.Semester ?? planGuessSem,
                        Year = chosen.Year,
                        ControlType = chosen.ControlType,
                        RequiresManualValidation = false
                    });
                    continue;
                }

                if (relaxed.Count > 1)
                {
                    statRelaxedAmbiguous++;
                    foreach (var t in relaxed)
                        pool.Remove(t);
                    logger.LogWarning(
                        "План: PlanOrder={Order} «{Name}» — неоднозначное relaxed-сопоставление ({Count} строк карточки), оценка не проставляется.",
                        pe.PlanOrder, planName, relaxed.Count);
                    output.Add(new StudentDisciplineResult
                    {
                        PlanBucket = pe.PlanBucket,
                        DisciplineName = planName,
                        PlanOrder = pe.PlanOrder,
                        AudHours = pe.TotalAudHours,
                        CreditUnits = totalCredits,
                        Semester = planGuessSem,
                        RequiresManualValidation = true
                    });
                    continue;
                }

                var truncatedPref = pool
                    .Where(c => c.DisciplineName != null &&
                        (IsTruncatedPlanCardNamePrefixMatch(planName, c.DisciplineName) ||
                         IsTruncatedPlanCardNamePrefixMatch(planName, StripSyntheticCourseTopicFromCard(c.DisciplineName))))
                    .ToList();

                if (truncatedPref.Count == 1)
                {
                    statTruncated++;
                    var chosen = truncatedPref[0];
                    pool.Remove(chosen);
                    logger.LogDebug(
                        "План↔карточка: PlanOrder={Order} «{Name}» — совпадение по префиксу усечённого названия с «{CardName}»",
                        pe.PlanOrder, planName, chosen.DisciplineName);
                    output.Add(new StudentDisciplineResult
                    {
                        PlanBucket = pe.PlanBucket,
                        DisciplineName = planName,
                        PlanOrder = pe.PlanOrder,
                        AudHours = pe.TotalAudHours,
                        CreditUnits = totalCredits,
                        Score = chosen.Score,
                        Semester = chosen.Semester ?? planGuessSem,
                        Year = chosen.Year,
                        ControlType = chosen.ControlType,
                        RequiresManualValidation = false
                    });
                    continue;
                }

                if (truncatedPref.Count > 1)
                {
                    statTruncatedAmbiguous++;
                    foreach (var t in truncatedPref)
                        pool.Remove(t);
                    logger.LogWarning(
                        "План: PlanOrder={Order} «{Name}» — неоднозначное сопоставление по префиксу усечённого названия ({Count} строк карточки), оценка не проставляется.",
                        pe.PlanOrder, planName, truncatedPref.Count);
                    output.Add(new StudentDisciplineResult
                    {
                        PlanBucket = pe.PlanBucket,
                        DisciplineName = planName,
                        PlanOrder = pe.PlanOrder,
                        AudHours = pe.TotalAudHours,
                        CreditUnits = totalCredits,
                        Semester = planGuessSem,
                        RequiresManualValidation = true
                    });
                    continue;
                }

                var loosenedPractice = pool
                    .Where(c => c.DisciplineName != null &&
                        LoosePlanPracticeNameEquals(planName, c.DisciplineName))
                    .ToList();

                if (loosenedPractice.Count == 1)
                {
                    statLoosenedPractice++;
                    var chosen = loosenedPractice[0];
                    pool.Remove(chosen);
                    logger.LogDebug(
                        "План↔карточка: PlanOrder={Order} «{Name}» — совпадение loosenedPractice с «{CardName}»",
                        pe.PlanOrder, planName, chosen.DisciplineName);
                    output.Add(new StudentDisciplineResult
                    {
                        PlanBucket = pe.PlanBucket,
                        DisciplineName = planName,
                        PlanOrder = pe.PlanOrder,
                        AudHours = pe.TotalAudHours,
                        CreditUnits = totalCredits,
                        Score = chosen.Score,
                        Semester = chosen.Semester ?? planGuessSem,
                        Year = chosen.Year,
                        ControlType = chosen.ControlType,
                        RequiresManualValidation = false
                    });
                    continue;
                }

                if (loosenedPractice.Count > 1)
                {
                    statLoosenedPracticeAmbiguous++;
                    foreach (var t in loosenedPractice)
                        pool.Remove(t);
                    logger.LogWarning(
                        "План: PlanOrder={Order} «{Name}» — неоднозначное сопоставление усечённой практики ({Count} строк карточки), оценка не проставляется.",
                        pe.PlanOrder, planName, loosenedPractice.Count);
                    output.Add(new StudentDisciplineResult
                    {
                        PlanBucket = pe.PlanBucket,
                        DisciplineName = planName,
                        PlanOrder = pe.PlanOrder,
                        AudHours = pe.TotalAudHours,
                        CreditUnits = totalCredits,
                        Semester = planGuessSem,
                        RequiresManualValidation = true
                    });
                    continue;
                }

                if (PlanNameLooksLikeCourseWorkPlanRow(planName))
                {
                    var courseCandidates = pool
                        .Where(c =>
                            c.ControlType != null &&
                            c.ControlType.IndexOf("курсов", StringComparison.OrdinalIgnoreCase) >= 0)
                        .ToList();

                    var narrowed = courseCandidates;
                    if (narrowed.Count > 1)
                    {
                        var nonPlaceholder = narrowed.Where(c => !CardCourseWorkLooksLikePlaceholderTopic(c)).ToList();
                        if (nonPlaceholder.Count == 1)
                            narrowed = nonPlaceholder;
                    }
                    if (narrowed.Count > 1 && planGuessSem.HasValue)
                    {
                        var bySem = narrowed.Where(c => c.Semester == planGuessSem).ToList();
                        if (bySem.Count == 1)
                            narrowed = bySem;
                    }

                    if (narrowed.Count == 1)
                    {
                        statCourseWorkBySemester++;
                        var chosen = narrowed[0];
                        pool.Remove(chosen);
                        logger.LogDebug(
                            "План↔карточка: PlanOrder={Order} «{Name}» — курсовая по слоту плана ↔ «{CardName}» (контроль={Ctl}, сем={Sem})",
                            pe.PlanOrder, planName, chosen.DisciplineName, chosen.ControlType, chosen.Semester);
                        output.Add(new StudentDisciplineResult
                        {
                            PlanBucket = pe.PlanBucket,
                            DisciplineName = planName,
                            PlanOrder = pe.PlanOrder,
                            AudHours = pe.TotalAudHours,
                            CreditUnits = totalCredits,
                            Score = chosen.Score,
                            Semester = chosen.Semester ?? planGuessSem,
                            Year = chosen.Year,
                            ControlType = chosen.ControlType,
                            RequiresManualValidation = false
                        });
                        continue;
                    }

                    if (narrowed.Count > 1)
                    {
                        logger.LogDebug(
                            "План↔карточка: PlanOrder={Order} «{Name}» — слот курсовой неоднозначен ({Count} кандидатов в карточке), идём дальше по стратегиям.",
                            pe.PlanOrder, planName, narrowed.Count);
                    }
                }

                var fuzzy = pool
                    .Where(c => c.DisciplineName != null &&
                        DisciplineNameSimilarityRatioForMatch(planName, c.DisciplineName) >= PlanCardFuzzySimilarityThreshold)
                    .ToList();

                if (fuzzy.Count == 1)
                {
                    statFuzzy++;
                    var chosen = fuzzy[0];
                    pool.Remove(chosen);
                    logger.LogDebug(
                        "План↔карточка: PlanOrder={Order} «{Name}» — fuzzy-сопоставление с «{CardName}» (порог сходства)",
                        pe.PlanOrder, planName, chosen.DisciplineName);
                    output.Add(new StudentDisciplineResult
                    {
                        PlanBucket = pe.PlanBucket,
                        DisciplineName = planName,
                        PlanOrder = pe.PlanOrder,
                        AudHours = pe.TotalAudHours,
                        CreditUnits = totalCredits,
                        Score = chosen.Score,
                        Semester = chosen.Semester ?? planGuessSem,
                        Year = chosen.Year,
                        ControlType = chosen.ControlType,
                        RequiresManualValidation = false
                    });
                    continue;
                }

                if (fuzzy.Count > 1)
                {
                    statFuzzyAmbiguous++;
                    foreach (var t in fuzzy)
                        pool.Remove(t);
                    logger.LogWarning(
                        "План: PlanOrder={Order} «{Name}» — неоднозначное fuzzy-сопоставление ({Count} кандидатов), оценка не проставляется.",
                        pe.PlanOrder, planName, fuzzy.Count);
                    output.Add(new StudentDisciplineResult
                    {
                        PlanBucket = pe.PlanBucket,
                        DisciplineName = planName,
                        PlanOrder = pe.PlanOrder,
                        AudHours = pe.TotalAudHours,
                        CreditUnits = totalCredits,
                        Semester = planGuessSem,
                        RequiresManualValidation = true
                    });
                    continue;
                }

                statPlanRowNoCard++;
                var topCand = pool
                    .Where(c => c.DisciplineName != null)
                    .Select(c => (
                        Card: c,
                        Sim: DisciplineNameSimilarityRatioForMatch(planName, c.DisciplineName)))
                    .OrderByDescending(t => t.Sim)
                    .Take(3)
                    .ToList();
                if (topCand.Count > 0)
                {
                    logger.LogDebug(
                        "План↔карточка: PlanOrder={Order} «{Name}» — нет соответствия; топ-{N} из пула карточки по похожести: {Top}",
                        pe.PlanOrder, planName, topCand.Count,
                        string.Join(" | ",
                            topCand.Select(t =>
                                $"«{t.Card.DisciplineName}» sim={t.Sim:0.###} сем={t.Card.Semester} тип={t.Card.ControlType}")));
                }
                logger.LogWarning("План: дисциплина PlanOrder={Order} «{Name}» — нет соответствия в карточке.", pe.PlanOrder, planName);
                output.Add(new StudentDisciplineResult
                {
                    PlanBucket = pe.PlanBucket,
                    DisciplineName = planName,
                    PlanOrder = pe.PlanOrder,
                    AudHours = pe.TotalAudHours,
                    CreditUnits = totalCredits,
                    Semester = planGuessSem,
                    RequiresManualValidation = true
                });
            }

            var statLeftoverNoise = 0;
            var statLeftoverCardOnly = 0;

            foreach (var left in pool)
            {
                if (SkipLeftoverDisciplineAfterPlan(left.DisciplineName))
                {
                    statLeftoverNoise++;
                    logger.LogWarning(
                        "Карточка: строка не сопоставлена с планом и отброшена как шум/служебная: «{Name}», нормДляСравнения={NormCt}, сем={Sem}",
                        left.DisciplineName ?? "(пусто)",
                        NormalizeDisciplineNameForComparison(left.DisciplineName),
                        left.Semester);
                    continue;
                }

                statLeftoverCardOnly++;
                logger.LogDebug(
                    "Карточка: хвост — «{Name}», темаИзСинтетики={Topic}, тип={Ctl}, сем={Sem}, оценка={Sc}",
                    left.DisciplineName ?? "(пусто)",
                    StripSyntheticCourseTopicFromCard(left.DisciplineName),
                    left.ControlType, left.Semester, left.Score);
                logger.LogWarning("Карточка: дисциплина «{Name}» не сопоставлена со строкой плана, будет отдельный блок в приложении.", left.DisciplineName ?? "(пусто)");
                output.Add(new StudentDisciplineResult
                {
                    PlanBucket = left.PlanBucket,
                    DisciplineName = left.DisciplineName,
                    Score = left.Score,
                    Semester = left.Semester,
                    Year = left.Year,
                    ControlType = left.ControlType,
                    AudHours = null,
                    CreditUnits = null,
                    PlanOrder = null,
                    RequiresManualValidation = true,
                    IsCardOnlyUnmatchedPlan = true
                });
            }

            logger.LogInformation(
                "План↔карточка: итог — строк результата={Out}; по стратегиям: exact={Ex}, exactНеодн={ExM}, relaxed={Rl}, relaxedНеодн={RlM}, префикс={Tr}, префиксНеодн={TrM}, практикаОслабл={Pr}, практикаНеодн={PrM}, курсоваяСлот={Crs}, fuzzy={Fz}, fuzzyНеодн={FzM}; строк плана без оценки из карточки={NoCard}; хвост карточки: отброшеноШум={Noise}, толькоКарточка={CardOnly}",
                output.Count,
                statExact, statExactAmbiguous, statRelaxed, statRelaxedAmbiguous,
                statTruncated, statTruncatedAmbiguous, statLoosenedPractice, statLoosenedPracticeAmbiguous,
                statCourseWorkBySemester,
                statFuzzy, statFuzzyAmbiguous, statPlanRowNoCard,
                statLeftoverNoise, statLeftoverCardOnly);

            return output;
        }

        public async Task<IEnumerable<Student>> GetStudentsJournal(IFormFile journal, List<Student> students)
        {
            var missingInJournal = new List<Student>();
            using (var stream = new MemoryStream())
            {
                await journal.CopyToAsync(stream);
                using (var package = new ExcelPackage(stream))
                {
                    var worksheet = package.Workbook.Worksheets[0] ?? throw new InvalidOperationException("Загруженный файл не содержит листов");

                    var columnCount = worksheet.Dimension.Columns;
                    var rowCount = worksheet.Dimension.Rows;

                    var headers = new List<string>();
                    for (int col = 1; col <= columnCount; col++)
                    {
                        headers.Add(worksheet.Cells[1, col].Text);
                    }

                    foreach (var student in students)
                    {
                        bool studentMatchedInFile = false;
                        string fio = $"{student.Surname}{student.Name}{student.Patronymic}".ToLower();

                        for (int row = 4; row <= rowCount; row++)
                        {
                            bool isCurrentStudent = false;
                            for (int col = 1; col <= columnCount; col++)
                            {
                                var header = headers[col - 1];
                                var cellValue = worksheet.Cells[row, col].Value ?? "";

                                if (header.ToLower() == "фио студента" || header.ToLower() == "фио обучающегося" || header.ToLower() == "фио")
                                {
                                    string cellfio = cellValue.ToString().ToLower().Replace(" ", "");
                                    if (cellfio == fio)
                                    {
                                        isCurrentStudent = true;
                                        studentMatchedInFile = true;
                                        break;
                                    }
                                }
                            }
                            for (int col = 1; col <= columnCount; col++)
                            {
                                var header = headers[col - 1];
                                var cellValue = worksheet.Cells[row, col].Value ?? "";

                                logger.LogInformation($"Работа с ячейкой: [{col},{row}]; столбец {header}");

                                switch (header.ToLower())
                                {
                                    case "№ студ.билета и зачетной книжки":
                                    case "№ студ.билета":
                                    case "№ зачетной книжки":
                                    case "№ зачетки":
                                    case "номер зачетки":
                                        if (isCurrentStudent)
                                            student.GradeBook = cellValue.ToString();
                                        break;
                                    case "№ группы":
                                    case "номер группы":
                                        if (isCurrentStudent)
                                            student.GroupName = cellValue.ToString();
                                        break;
                                }
                            }
                        }
                        if (!studentMatchedInFile)
                            missingInJournal.Add(student);
                    }
                }
            }
            if (missingInJournal.Count > 0)
            {
                throw new StudentImportMismatchException(
                    "журнал зачётных книжек (3-й файл импорта, назначение групп)",
                    journal.FileName,
                    missingInJournal.Select(FormatStudentDisplayName).ToList());
            }

            return students;
        }

        public async Task<IEnumerable<Student>> GetStudentsLD(IFormFile ld)
        {
            var students = new List<Student>();
            using (var stream = new MemoryStream())
            {
                await ld.CopyToAsync(stream);
                using (var package = new ExcelPackage(stream))
                {
                    var worksheet = package.Workbook.Worksheets[0] ?? throw new InvalidOperationException("Загруженный файл не содержит листов");

                    var columnCount = worksheet.Dimension.Columns;
                    var rowCount = worksheet.Dimension.Rows;

                    var headers = new List<string>();
                    for (int col = 1; col <= columnCount; col++)
                    {
                        headers.Add(worksheet.Cells[1, col].Text);
                    }

                    for (int row = 2; row <= rowCount; row++)
                    {
                        var student = new Student();
                        string courseOfTraining = string.Empty;

                        // Фиксированные данные
                        student.Education = "высшее"; // других видов нет
                        student.EducationForm = "очная"; // других форм нет
                        student.Faculty = "Информационных технологий"; // Т.к. делаем прогу для нашего факультета, факультет будет такой. Пока взять альтернативные варианты неоткуда
                        student.Course = "Прикладная информатика в психологии"; // нужна таблица с соответствиями специализаций с направлениями подготовки
                        student.EducationBase = "бюджетная (ФБ)";// других основ нет
                        student.EducationRelationForm = "договор об образовании на обучение";// платных услуг нет
                        student.EducationTime = 4; // Пока ставль 4, т.к. неоткуда брать данные
                        student.LivingInDormitory = false; // Общежитие по умолчанию нет
                        student.MilitaryService = false; // Служба в армии по умолчанию нет
                        student.MaritalStatus = false; // Отношения по умолчанию отсутствуют


                        for (int col = 1; col <= columnCount; col++)
                        {
                            var header = headers[col - 1];
                            var cellValue = worksheet.Cells[row, col].Value ?? "";

                            string indexPattern = @"\b\d{6}\b";
                            string cityPattern = @"\b(?:(?<city>[\w\s]+?)\s+[гсхдп]\,?|(?<type>[гсхдп])\.?\s*(?<city>[\w\s]+?))\b";
                            string streetPattern = @",?\s*([^,]+?)\s*,\s*(?:дом|д)\.";
                            string housingTypePattern = @"(?:^|,)\s*(к(?:\.|орпус)?|стр(?:\.|оение)?)(?=\s*\w|$)(?!\w)";
                            string housingPattern = @"(?:^|,)\s*(?<type>к(?:\.|орпус)?|стр(?:\.|оение)?)\s*(?<number>[\w\-]*\d[\w\-]*)\b";
                            string apartementPattern = @"(?:^|\s)(?:кв\.?|квартира\.?)\s*(\w+)\b";
                            string numConcursPattern = @"(\d{2}\.\d{2}\.\d{2})";
                            string regionPattern = @"(?:^|,)\s*(?<region>[\w\s-]+?)\s*(?<type>обл\.?|кр\.?|край|автономная\s+область|авт\.?\s*обл\.?|АО)\b";

                            logger.LogInformation($"Работа с ячейкой: [{col},{row}]; столбец {header}");


                            switch (header.ToLower())
                            {
                                case "фио":
                                    if (cellValue == "") return students;
                                    break;
                            }

                            switch (header.ToLower())
                            {
                                case "фио":
                                    string pattern = @"\S+";
                                    MatchCollection fio = Regex.Matches(cellValue.ToString().ToLower(), pattern);
                                    string name = $"{fio[1].ToString().Substring(0, 1).ToUpper()}{fio[1].ToString().Substring(1)}";
                                    student.Name = name;
                                    string surname = $"{fio[0].ToString().Substring(0, 1).ToUpper()}{fio[0].ToString().Substring(1)}";
                                    student.Surname = surname;
                                    if (fio.Count > 2)
                                    {
                                        string patronymic = $"{fio[2].ToString().Substring(0, 1).ToUpper()}{fio[2].ToString().Substring(1)}";
                                        student.Patronymic = patronymic;
                                    }
                                    else student.Patronymic = "";
                                    break;
                                case "пол":
                                    if (cellValue.ToString().ToLower() == "мужской") student.Gender = true;
                                    else student.Gender = false;
                                    break;
                                case "дата рождения":
                                    if (cellValue is DateTime birDate)
                                    {
                                        student.BirthdayDate = DateOnly.FromDateTime(birDate);
                                    }
                                    else
                                    {
                                        string dateStr = cellValue.ToString().Trim();
                                        if (DateTime.TryParseExact(
                                            dateStr,
                                            "dd.MM.yyyy",
                                            CultureInfo.InvariantCulture,
                                            DateTimeStyles.None,
                                            out DateTime parsedDate))
                                        {
                                            student.BirthdayDate = DateOnly.FromDateTime(parsedDate);
                                        }
                                        else throw new FormatException($"Не удалось распознать дату: {dateStr}");
                                    }
                                    break;
                                case "место рождения":
                                    student.BirthdayPlace = cellValue.ToString();
                                    break;
                                case "телефон":
                                    student.PhoneNumber = cellValue.ToString();
                                    break;
                                case "e-mail":
                                case "почта":
                                    student.Email = cellValue.ToString();
                                    break;
                                case "серия документа удостоверяющего личность":
                                    student.PassportSerial = cellValue.ToString();
                                    break;
                                case "номер документа удостоверяющего личность":
                                    student.PassportNumber = cellValue.ToString();
                                    break;
                                case "овддокумента удостоверяющего личность":
                                    student.PassportIssuancePlace = cellValue.ToString();
                                    break;
                                case "код подразделения документа удостоверяющего личность":
                                case "код подразделения":// Здесь нужны уточнения, как называется поле
                                    student.PassportIssuanceCode = cellValue.ToString();
                                    break;
                                case "дата выдачи паспорта": // Здесь нужны уточнения, как называется поле и существует ли вообще
                                    if (cellValue is DateTime pasDate)
                                    {
                                        student.PassportIssuanceDate = DateOnly.FromDateTime(pasDate);
                                    }
                                    else
                                    {
                                        string dateStr = cellValue.ToString().Trim();
                                        if (DateTime.TryParseExact(
                                            dateStr,
                                            "dd.MM.yyyy",
                                            CultureInfo.InvariantCulture,
                                            DateTimeStyles.None,
                                            out DateTime parsedDate))
                                        {
                                            student.PassportIssuanceDate = DateOnly.FromDateTime(parsedDate);
                                        }
                                        else throw new FormatException($"Не удалось распознать дату: {dateStr}");
                                    }
                                    break;
                                case "гражданство":
                                    student.Citizenship = cellValue.ToString();
                                    break;
                                case "предмет1":
                                    if (cellValue != "")
                                        student.GiaExam1Score = short.Parse(cellValue.ToString());
                                    break;
                                case "предмет2":
                                    if (cellValue != "")
                                        student.GiaExam2Score = short.Parse(cellValue.ToString());
                                    break;
                                case "предмет3":
                                    if (cellValue != "")
                                        student.GiaExam3Score = short.Parse(cellValue.ToString());
                                    break;
                                case "сумма баллов за инд.дост.(конкурсные)":
                                case "сумма баллов за инд.дост.":
                                    if (cellValue != "")
                                        student.BonusScores = short.Parse(cellValue.ToString());
                                    break;
                                case "адрес по прописке":
                                    if (cellValue.ToString() != "")
                                    {
                                        var registrationAddress = cellValue.ToString();
                                        student.AddressRegistrationIndex = Regex.Match(registrationAddress, indexPattern).ToString();
                                        student.AddressRegistrationCity = Regex.Match(registrationAddress, cityPattern).Groups[1].ToString();
                                        ApplyAddressSettlementType(
                                            student,
                                            ParseAddressSettlementTypeAbbrev(registrationAddress),
                                            isRegistration: true);
                                        student.AddressRegistrationHouse = ParseAddressHouseNumber(registrationAddress);
                                        student.AddressRegistrationStreet = Regex.Match(registrationAddress, streetPattern).Groups[1].ToString().Trim();
                                        string housingMatch = Regex.Match(registrationAddress, housingTypePattern).Groups[1].ToString().Trim();
                                        if (housingMatch == "к" || housingMatch == "к." || housingMatch == "корпус" || housingMatch == "корпус.")
                                            student.AddressRegistrationHousingType = "корпус";
                                        else
                                            if (housingMatch == "стр" || housingMatch == "стр." || housingMatch == "строение" || housingMatch == "строение.")
                                            student.AddressRegistrationHousingType = "строение";
                                        student.AddressRegistrationHousing = Regex.Match(registrationAddress, housingPattern).Groups[2].ToString().Trim();
                                        student.AddressRegistrationApartment = Regex.Match(registrationAddress, apartementPattern).Groups[1].ToString().Trim();
                                        student.AddressRegistrationOblKrayAvtobl = Regex.Match(registrationAddress, regionPattern).Groups[1].ToString().Trim();

                                    }
                                    break;

                                case "адрес проживания":
                                    if (cellValue.ToString() != "")
                                    {
                                        var residentialAddress = cellValue.ToString();
                                        student.AddressResidentialIndex = Regex.Match(residentialAddress, indexPattern).ToString();
                                        student.AddressResidentialCity = Regex.Match(residentialAddress, cityPattern).Groups[1].ToString();
                                        ApplyAddressSettlementType(
                                            student,
                                            ParseAddressSettlementTypeAbbrev(residentialAddress),
                                            isRegistration: false);
                                        student.AddressResidentialHouse = ParseAddressHouseNumber(residentialAddress);
                                        student.AddressResidentialStreet = Regex.Match(residentialAddress, streetPattern).Groups[1].ToString().Trim();
                                        string housingMatch = Regex.Match(residentialAddress, housingTypePattern).Groups[1].ToString().Trim();
                                        if (housingMatch == "к" || housingMatch == "к." || housingMatch == "корпус" || housingMatch == "корпус.")
                                            student.AddressResidentialHousingType = "корпус";
                                        else
                                            if (housingMatch == "стр" || housingMatch == "стр." || housingMatch == "строение" || housingMatch == "строение.")
                                            student.AddressResidentialHousingType = "строение";
                                        student.AddressResidentialHousing = Regex.Match(residentialAddress, housingPattern).Groups[2].ToString().Trim();
                                        student.AddressResidentialApartment = Regex.Match(residentialAddress, apartementPattern).Groups[1].ToString().Trim();
                                        student.AddressResidentialOblKrayAvtobl = Regex.Match(residentialAddress, regionPattern).Groups[1].ToString().Trim();

                                    }
                                    break;
                                case "конкурсная группа":
                                    courseOfTraining = Regex.Match(cellValue.ToString(), numConcursPattern).Groups[1].ToString().Trim();
                                    break;
                                case "направление\\специальность":
                                    student.CourseOfTraining = $"{courseOfTraining} {cellValue.ToString()}";
                                    break;
                                case "серия документа об образовании":
                                    student.EducationReceivedSerial = cellValue.ToString();
                                    break;
                                case "номер документа об образовании":
                                    student.EducationReceivedNum = cellValue.ToString();
                                    break;
                                case "дата выдачи":
                                    if (cellValue is DateTime exDate) student.EducationReceivedDate = DateOnly.FromDateTime(exDate);
                                    else
                                    {
                                        string dateStr = cellValue.ToString().Trim();
                                        if (DateTime.TryParseExact(
                                            dateStr,
                                            "dd.MM.yyyy",
                                            CultureInfo.InvariantCulture,
                                            DateTimeStyles.None,
                                            out DateTime parsedDate))
                                        {
                                            student.EducationReceivedDate = DateOnly.FromDateTime(parsedDate);
                                        }
                                        else throw new FormatException($"Не удалось распознать дату: {dateStr}");
                                    }
                                    break;
                                case "год завершения":
                                    student.EducationReceivedEndYear = short.Parse(cellValue.ToString());
                                    break;
                                case "тип документа об образовании":
                                    if (cellValue.ToString().ToLower() == "аттестат")
                                        student.EducationReceived = "среднее общее образование";
                                    else if (cellValue.ToString().ToLower() == "диплом")
                                        student.EducationReceived = "среднее профессиональное образование";
                                    break;
                                case "образовательное учреждение":
                                    student.OOName = cellValue.ToString();
                                    break;
                            }
                            if ((header.ToLower() == "адрес проживания") && (cellValue.ToString() == ""))
                            {
                                student.AddressResidentialCity = student.AddressRegistrationCity;
                                student.AddressResidentialApartment = student.AddressRegistrationApartment;
                                student.AddressResidentialDistrict = student.AddressRegistrationDistrict;
                                student.AddressResidentialHouse = student.AddressRegistrationHouse;
                                student.AddressResidentialHousing = student.AddressRegistrationHousing;
                                student.AddressResidentialHousingType = student.AddressRegistrationHousingType;
                                student.AddressResidentialIndex = student.AddressRegistrationIndex;
                                student.AddressResidentialOblKrayAvtobl = student.AddressRegistrationOblKrayAvtobl;
                                student.AddressResidentialStreet = student.AddressRegistrationStreet;
                                student.AddressResidentialType = student.AddressRegistrationType;
                            }
                            if (student.EducationReceivedEndYear == null)
                            {
                                if (student.EducationReceivedDate.HasValue)
                                {
                                    student.EducationReceivedEndYear = (short)student.EducationReceivedDate.Value.Year;
                                }
                            }
                        }
                        students.Add(student);
                    }
                }

            }
            return students;
        }

        // Функция извлечения данных из ведомости
        public async Task<IEnumerable<Student>> GetStudentsStatement(IFormFile statement, List<Student> students)
        {
            // Открытие потока обмеена данных
            using (var stream = new MemoryStream())
            {
                await statement.CopyToAsync(stream);

                // Распаковка файла Excel
                using (var package = new ExcelPackage(stream))
                {
                    // Заголовки для исключения из списка заголовков
                    var excludedHeaders = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                    {
                        "№ п/п",
                        "фио студента",
                        "фио обучающегося",
                        "фио",
                        "обучается за счёт средств бюджета",
                        "обучаюется за счёт средств бюджета",
                        "обучается за счёт средств бюджета*",
                        "обучаюется за счёт средств бюджета*",
                        "средний балл",
                        "примечание"
                    };

                    bool IsExcludedHeader(string? header)
                    {
                        if (string.IsNullOrWhiteSpace(header))
                            return true;
                        return excludedHeaders.Contains(header.Trim());
                    }

                    bool TryGetNumericScore(object? value, out double score)
                    {
                        score = default;
                        if (value == null)
                            return false;

                        if (value is double d)
                        {
                            score = d;
                            return true;
                        }

                        if (value is float f)
                        {
                            score = f;
                            return true;
                        }

                        if (value is decimal dec)
                        {
                            score = (double)dec;
                            return true;
                        }

                        if (value is int i)
                        {
                            score = i;
                            return true;
                        }

                        if (value is long l)
                        {
                            score = l;
                            return true;
                        }

                        var str = value.ToString()?.Trim();
                        if (string.IsNullOrWhiteSpace(str))
                            return false;

                        if (double.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out score))
                            return true;


                        var normalized = str.Replace(" ", string.Empty).Replace(",", ".");
                        return double.TryParse(normalized, NumberStyles.Any, CultureInfo.InvariantCulture, out score);
                    }

                    static string NormalizeStatementFio(string? value)
                    {
                        if (string.IsNullOrWhiteSpace(value))
                            return string.Empty;
                        return value.ToLower().Replace(" ", string.Empty).Replace(".", string.Empty);
                    }

                    static bool IsFioHeader(string? header)
                    {
                        if (string.IsNullOrWhiteSpace(header))
                            return false;
                        return header.ToLower().Contains("фио");
                    }

                    static string NormalizeHeaderText(string? value)
                    {
                        if (string.IsNullOrWhiteSpace(value))
                            return string.Empty;

                        var collapsed = Regex.Replace(value, @"\s+", " ").Trim();
                        collapsed = Regex.Replace(collapsed, @"\(\s*зач[её]т\s*с\s*оценкой\s*\)", string.Empty, RegexOptions.IgnoreCase);
                        collapsed = Regex.Replace(collapsed, @"\(\s*Итоговая\s+контрольная\s+работа\s*\)", string.Empty, RegexOptions.IgnoreCase);
                        return Regex.Replace(collapsed, @"\s+", " ").Trim();
                    }

                    // Проверка файла на наличие листов
                    if (package.Workbook.Worksheets.Count == 0)
                        throw new InvalidOperationException("Загруженный файл не содержит листов");

                    foreach (var worksheet in package.Workbook.Worksheets)
                    {
                        if (worksheet?.Dimension == null)
                            continue;

                        // Базовые данные листов
                        var columnCount = worksheet.Dimension.Columns;
                        var rowCount = worksheet.Dimension.Rows;

                        string GetMergedText(int row, int col)
                        {
                            var mergedAddress = worksheet.MergedCells[row, col];
                            if (!string.IsNullOrWhiteSpace(mergedAddress))
                                return worksheet.Cells[mergedAddress].First().Text;
                            return worksheet.Cells[row, col].Text;
                        }

                        static bool TryExtractFirstInt(string? value, out int number)
                        {
                            number = default;
                            if (string.IsNullOrWhiteSpace(value))
                                return false;

                            var match = Regex.Match(value, @"\d+");
                            if (!match.Success)
                                return false;

                            return int.TryParse(match.Value, out number);
                        }

                        short? sheetSemester = null;
                        {
                            var sessionRowText = NormalizeHeaderText(GetMergedText(3, 1)).ToLower();
                            var isAutumnWinter = sessionRowText.Contains("осенне-зимняя") || sessionRowText.Contains("осенне зимняя");
                            var isSpringSummer = sessionRowText.Contains("весенне-летняя") || sessionRowText.Contains("весенне летняя");

                            // Номер курса — только из 4-й строки (не из года/даты в шапке)
                            int? courseNum = TryExtractCourseFromStatementRow(4, columnCount, GetMergedText, out var parsedCourse)
                                ? parsedCourse
                                : null;

                            if (courseNum.HasValue && courseNum.Value > 0)
                            {
                                var sem = courseNum.Value * 2;
                                if (isAutumnWinter)
                                    sem -= 1;
                                sheetSemester = (short)sem;
                            }
                        }

                        short? sheetYear = null;
                        {
                            var headerRowText = NormalizeHeaderText(GetMergedText(2, 1));
                            if (TryExtractFirstInt(headerRowText, out var yearCandidate) && yearCandidate >= 1900 && yearCandidate <= 2100)
                            {
                                // TryExtractFirstInt найдёт первое число; в строке может быть № ведомости.
                                // Поэтому ищем именно год как 4 цифры.
                                var yearMatch = Regex.Match(headerRowText, @"\b(19\d{2}|20\d{2}|2100)\b");
                                if (yearMatch.Success && short.TryParse(yearMatch.Value, out var parsedYear))
                                    sheetYear = parsedYear;
                            }
                            else
                            {
                                var yearMatch = Regex.Match(headerRowText, @"\b(19\d{2}|20\d{2}|2100)\b");
                                if (yearMatch.Success && short.TryParse(yearMatch.Value, out var parsedYear))
                                    sheetYear = parsedYear;
                            }
                        }

                        var headers = new List<string>();
                        var topHeaders = new List<string>();

                        // Извлечение заголовков из объединённых 6-7 строк
                        for (int col = 1; col <= columnCount; col++)
                        {
                            var top = NormalizeHeaderText(GetMergedText(6, col));
                            var bottom = NormalizeHeaderText(GetMergedText(7, col));

                            topHeaders.Add(top);

                            string combined;
                            if (string.IsNullOrWhiteSpace(top))
                                combined = bottom;
                            else if (string.IsNullOrWhiteSpace(bottom))
                                combined = top;
                            else
                                combined = $"{top} {bottom}";

                            headers.Add(NormalizeHeaderText(combined));
                        }

                        var studentRows = new Dictionary<Student, int>();

                        foreach (var student in students)
                        {
                            // Фамилия и имя студента в таблице
                            var surname = student.Surname ?? string.Empty;
                            var nameInitial = string.IsNullOrWhiteSpace(student.Name) ? string.Empty : student.Name.Substring(0, 1);
                            var patronymicInitial = string.IsNullOrWhiteSpace(student.Patronymic) ? string.Empty : student.Patronymic.Substring(0, 1);

                            var keyFi = NormalizeStatementFio($"{surname}{nameInitial}");
                            var keyFip = NormalizeStatementFio($"{surname}{nameInitial}{patronymicInitial}");
                            // Перебор начиная с 8 строки
                            for (int row = 8; row <= rowCount; row++)
                            {
                                // Поиск студентов из списка
                                for (int col = 1; col <= columnCount; col++)
                                {
                                    var header = headers[col - 1];
                                    var cellValue = worksheet.Cells[row, col].Value ?? "";

                                    if (!IsFioHeader(header))
                                        continue;

                                    var cellfi = NormalizeStatementFio(cellValue.ToString());
                                    if (cellfi.Contains(keyFip) || cellfi.Contains(keyFi))
                                    {
                                        studentRows[student] = row;
                                        break;
                                    }
                                    if (studentRows.ContainsKey(student))
                                        break;
                                }
                            }
                        }

                        // Обработка данных для студентов
                        foreach (var student in students)
                        {
                            if (!studentRows.TryGetValue(student, out int studentRow))
                                continue;

                            for (int col = 1; col <= columnCount; col++)
                            {
                                var header = headers[col - 1];
                                var headerLower = header.ToLower();

                                if (IsExcludedHeader(header))
                                    continue;

                                var cellValue = worksheet.Cells[studentRow, col].Value;
                                if (!TryGetNumericScore(cellValue, out var score))
                                {
                                    var cellText = cellValue?.ToString()?.Trim() ?? string.Empty;
                                    var normalized = cellText.ToLower().Replace(".", string.Empty).Replace(" ", string.Empty);
                                    if (normalized.Contains("неявл"))
                                    {
                                        score = 0;
                                    }
                                    else
                                    {
                                        continue;
                                    }
                                }

                                var disciplineName = (topHeaders[col - 1] ?? "").Trim();
                                if (string.IsNullOrWhiteSpace(disciplineName))
                                    disciplineName = header.Trim();

                                disciplineName = NormalizeHeaderText(disciplineName);

                                var isCourseWork = disciplineName.ToLower().Contains("курсовая работа");
                                if (isCourseWork)
                                    disciplineName = string.Empty;

                                // Обработка практик: "Учебная практика, ..." или "Производственная практика, ..."
                                var isPractice = false;
                                var productionPracticeMatch = Regex.Match(disciplineName, @"^Производственная практика,", RegexOptions.IgnoreCase);
                                var studyPracticeMatch = Regex.Match(disciplineName, @"^Учебная практика,", RegexOptions.IgnoreCase);
                                
                                if (productionPracticeMatch.Success || studyPracticeMatch.Success)
                                {
                                    isPractice = true;
                                    // Для практик НЕ убирается префикс, название остается как есть
                                    // Но при сравнении будет использоваться NormalizeDisciplineNameForComparison
                                }

                                if (!isCourseWork && !isPractice)
                                {
                                    if (IsExcludedHeader(disciplineName) || excludedHeaders.Contains(disciplineName))
                                        continue;
                                }

                                logger.LogInformation($"Работа с ячейкой: [{col},{studentRow}]; столбец {headerLower}");

                                var existing = isCourseWork
                                    ? student.DisciplineResults.FirstOrDefault(x =>
                                        x.ControlType != null &&
                                        x.ControlType.Equals("курсовая", StringComparison.OrdinalIgnoreCase))
                                    : isPractice
                                        ? student.DisciplineResults.FirstOrDefault(x =>
                                            x.DisciplineName != null &&
                                            x.ControlType != null &&
                                            x.ControlType.Equals("практика", StringComparison.OrdinalIgnoreCase) &&
                                            NormalizeDisciplineNameForComparison(x.DisciplineName).Equals(NormalizeDisciplineNameForComparison(disciplineName), StringComparison.OrdinalIgnoreCase))
                                        : student.DisciplineResults.FirstOrDefault(x =>
                                            x.DisciplineName != null &&
                                            x.DisciplineName.Equals(disciplineName, StringComparison.OrdinalIgnoreCase));

                                string? controlType = isCourseWork ? "курсовая" : isPractice ? "практика" : null;

                                if (existing == null)
                                {
                                    student.DisciplineResults.Add(new StudentDisciplineResult
                                    {
                                        DisciplineName = string.IsNullOrWhiteSpace(disciplineName) ? null : disciplineName,
                                        Score = score.ToString(),
                                        Semester = sheetSemester,
                                        Year = sheetYear,
                                        ControlType = controlType
                                    });
                                }
                                else
                                {
                                    existing.Score = score.ToString();
                                    existing.Semester = sheetSemester;
                                    existing.Year = sheetYear;
                                    if (controlType != null)
                                        existing.ControlType = controlType;
                                }
                            }
                        }
                    }
                }
            }
            return students;
        }

        public async Task<DiplomaSupplementData> GetStudentCardAsync(IFormFile studentCard, IFormFile plan)
        {
            ArgumentNullException.ThrowIfNull(plan);

            var diplomaData = new DiplomaSupplementData();
            using (var stream = new MemoryStream())
            {
                await studentCard.CopyToAsync(stream);
                using (var package = new ExcelPackage(stream))
                {
                    // Важно: рассчитываем формулы, если это возможно, но для сломанных ссылок будем брать Text
                    package.Workbook.Calculate();

                    // Начальный скан документа (Лист 1)
                    var firstSheet = package.Workbook.Worksheets[0] ?? throw new InvalidOperationException("Загруженный файл не содержит листов");
                    var infoSheet = package.Workbook.Worksheets["ОбщСведения"] ?? firstSheet;
                    var courseText = infoSheet.Cells[78, 3].Text?.Trim();
                    if (!string.IsNullOrWhiteSpace(courseText))
                        diplomaData.CourseOfTraining = courseText;

                    var qualFromC110 = NormalizeSupplementOpopSheetText(infoSheet.Cells[110, 3].Text);
                    if (!string.IsNullOrWhiteSpace(qualFromC110))
                        diplomaData.SupplementOwnerQualification = Capitalize(qualFromC110);

                    var opopFromC79 = NormalizeSupplementOpopSheetText(infoSheet.Cells[79, 3].Text);
                    if (!string.IsNullOrWhiteSpace(opopFromC79))
                        diplomaData.SupplementAdditionalSheetOpopName = opopFromC79;
                    else
                    {
                        var fallbackOp = NormalizeSupplementOpopSheetText(courseText);
                        if (!string.IsNullOrWhiteSpace(fallbackOp))
                            diplomaData.SupplementAdditionalSheetOpopName = fallbackOp;
                    }

                    diplomaData.SupplementAdditionalSheetStudyFormLine =
                        BuildSupplementStudyFormLine(infoSheet.Cells[76, 8].Text);

                    diplomaData.TargetProgramCredits = await ReadTargetProgramCreditsFromPlanAsync(plan);
                    diplomaData.TargetContactHoursFromPlan = await ReadTargetContactHoursFromPlanBlockSummariesAsync(plan);
                    diplomaData.TargetGiaCreditsFromPlan = await ReadTargetGiaCreditsFromPlanAsync(plan);
                    diplomaData.TargetPracticeCreditsFromPlan = await ReadTargetPracticeCreditsFromPlanAsync(plan);

                    var enableFacultyParsing = configuration.GetValue("EnableFacultyParsing", false);
                    logger.LogInformation("EnableFacultyParsing={Enable}", enableFacultyParsing);

                    // 1. Парсинг основной инфо (как было)
                    var mappings = new Dictionary<(int row, int col), Action<string>>
                    {
                        [(3, 2)] = val => diplomaData.Surname = Capitalize(val),
                        [(4, 2)] = val => diplomaData.Name = Capitalize(val),
                        [(5, 2)] = val => diplomaData.Patronymic = Capitalize(val),
                        [(40, 4)] = val => diplomaData.EducationReceived = Capitalize(val)
                    };

                    foreach (var ((row, col), setter) in mappings)
                    {
                        var cell = firstSheet.Cells[row, col];
                        logger.LogDebug($"Значение ячейки [{row}, {col}]: {cell.Text}");
                        setter(cell.Text);
                    }

                    diplomaData.BirthdayDate = ReadSupplementSheetDateCell(firstSheet, 6, 5);
                    diplomaData.EducationReceivedDate = ReadSupplementSheetDateCell(firstSheet, 42, 8);

                    if (!diplomaData.BirthdayDate.HasValue)
                        logger.LogWarning("Карточка supplement: нет даты рождения (лист «{Sheet}», ожидалась ячейка около строки 6).", firstSheet.Name);
                    if (!diplomaData.EducationReceivedDate.HasValue)
                        logger.LogWarning("Карточка supplement: нет даты документа об образовании (ожидалась строка около 42).");

                    // 2. Парсинг набора оценок (семестры)
                    var (disciplineGrades, honors) = ParseDisciplineGradesFromCard(package.Workbook.Worksheets);

                    // 3. Парсинг финализации (Гос. экз и ВКР) с первого листа
                    var finalizationResults = ParseFinalizationData(firstSheet);

                    // 4. Список дисциплин от учебного плана + сопоставление с карточкой
                    var planEntries = await ParseStudyPlanForSupplementAsync(plan, enableFacultyParsing);
                    logger.LogInformation(
                        "Supplement: разбор ПланСвод завершён, записей плана={PlanN}; переход к сопоставлению с карточкой ({CardGrades} оценок).",
                        planEntries.Count, disciplineGrades.Count);
                    var fromPlan = BuildDisciplineResultsFromPlanAndCard(planEntries, disciplineGrades);
                    if (finalizationResults.Any())
                    {
                        fromPlan.AddRange(finalizationResults);
                        logger.LogInformation($"Добавлено {finalizationResults.Count} записей итоговой аттестации (Гос.экзамен/ВКР).");
                    }

                    diplomaData.DisciplineResults = fromPlan;
                    diplomaData.DiplomaWithHonors = honors;
                }
            }

            return diplomaData;
        }

        private static string? NormalizeSupplementOpopSheetText(string? raw)
        {
            var s = NormalizeDisciplineName(raw ?? "");
            if (string.IsNullOrWhiteSpace(s))
                return null;
            if (s.Length >= 2 && s.StartsWith('"') && s.EndsWith('"'))
                s = s.Substring(1, s.Length - 2).Trim();
            return string.IsNullOrWhiteSpace(s) ? null : s;
        }

        private static string? BuildSupplementStudyFormLine(string? rawFromCard)
        {
            if (string.IsNullOrWhiteSpace(rawFromCard))
                return null;
            var t = NormalizeDisciplineName(rawFromCard).ToLowerInvariant();
            if (t.Contains("заоч"))
                return "Форма обучения: заочное";
            if (t.Contains("очн"))
                return "Форма обучения: очная";
            return "Форма обучения: " + NormalizeDisciplineName(rawFromCard);
        }

        private string Capitalize(string val)
        {
            if (string.IsNullOrWhiteSpace(val)) return val;
            val = val.Trim();
            if (val.Length == 0) return val;
            return char.ToUpper(val[0]) + val.Substring(1);
        }

        /// <summary>
        /// Парсит блок финализации (строки 102-107 первого листа)
        /// </summary>
        private List<StudentDisciplineResult> ParseFinalizationData(ExcelWorksheet sheet)
        {
            var results = new List<StudentDisciplineResult>();

            // Проверяем наличие заголовка "Финализация" (не обязательно, но полезно для валидации)
            var header = sheet.Cells["A102"].Text;
            if (string.IsNullOrWhiteSpace(header))
            {
                logger.LogWarning("Не найден заголовок 'Финализация' в ячейке A102. Пытаемся считать данные по фиксированным координатам.");
            }

            // --- 1. Государственный экзамен (Строка 104) ---
            // Оценка: D104 (объединена D-E)
            // Дата: G104
            string stateExamScoreRaw = sheet.Cells["D104"].Text;

            // Если оценка есть (не пусто), добавляем запись
            if (!string.IsNullOrWhiteSpace(stateExamScoreRaw))
            {
                var date = ObjectToDateOnly(sheet.Cells["G104"].Value, false);

                var stateExam = new StudentDisciplineResult
                {
                    DisciplineName = "Государственный экзамен",
                    Score = MapGradeTo15Scale(stateExamScoreRaw),
                    ControlType = "экзамен", // или "государственный экзамен"
                    CreditUnits = 0, // Не указано в задаче, ставим 0
                    AudHours = 0,
                    Year = date is null ? null : (short)date.Value.Year,
                    Semester = null // Итоговая аттестация вне семестров
                };
                results.Add(stateExam);
                logger.LogDebug($"Добавлен гос. экзамен: {stateExam.Score} ({stateExamScoreRaw})");
            }

            // --- 2. Защита ВКР (Строка 105 + Тема 107) ---
            // Оценка: D105
            // Дата: G105
            // Тема: A107 (объединена A-J)
            string vkrScoreRaw = sheet.Cells["D105"].Text;

            if (!string.IsNullOrWhiteSpace(vkrScoreRaw))
            {
                var date = ObjectToDateOnly(sheet.Cells["G105"].Value, false);
                string topicRaw = sheet.Cells["A107"].Text;

                // Очистка темы от кавычек, если они есть в начале/конце, 
                // так как мы будем оборачивать её в формат.
                string topicClean = topicRaw?.Trim() ?? "";
                if (topicClean.StartsWith("\"") && topicClean.EndsWith("\"") && topicClean.Length > 1)
                {
                    topicClean = topicClean.Substring(1, topicClean.Length - 2);
                }

                if (string.IsNullOrWhiteSpace(topicClean))
                {
                    topicClean = "Тема выпускной квалификационной работы";
                }

                // Формируем имя ВКР. 
                // Вариант: "Выпускная квалификационная работа (бакалаврская работа). Тема: \"...\""
                // Чтобы экспорт мог корректно разбить это на строки.
                string vkrName = $"Выпускная квалификационная работа. Тема: \"{topicClean}\"";

                var vkr = new StudentDisciplineResult
                {
                    DisciplineName = vkrName,
                    Score = MapGradeTo15Scale(vkrScoreRaw),
                    ControlType = "защита вкр",
                    CreditUnits = 0,
                    AudHours = 0,
                    Year = date is null ? null : (short)date.Value.Year,
                    Semester = null
                };
                results.Add(vkr);
                logger.LogDebug($"Добавлена ВКР: {vkrName}, Оценка: {vkr.Score}");
            }

            return results;
        }

        private (List<StudentDisciplineResult> disciplineResults, bool diplomaWithHonors) ParseDisciplineGradesFromCard(ExcelWorksheets worksheets)
        {
            logger.LogInformation("Начало парсинга оценок из карточки студента");

            // Отбираем все листы в экселе с названиями "Ро..."
            var eduResWorksheets = worksheets.Where(w => w.Name.Contains("Ро", StringComparison.OrdinalIgnoreCase));

            logger.LogDebug($"Найдено листов с названиями 'Ро...': {eduResWorksheets.Count()}");

            if (!eduResWorksheets.Any())
            {
                logger.LogError("Не были найдены листы с названиями \"Ро...\"");
                throw new ArgumentException("Не были найдены листы с названиями \"Ро...\" с результатами обучения в карточке студента.");
            }

            eduResWorksheets = eduResWorksheets.Reverse();
            logger.LogTrace("Листы будут обрабатываться в обратном порядке для получения самых последних оценок");

            var allResults = new List<StudentDisciplineResult>();

            // Собираем все результаты
            foreach (var sheet in eduResWorksheets)
            {
                logger.LogDebug($"Обработка листа: {sheet.Name}");
                var sheetResults = ParseEducationResultsWorksheet(sheet);
                logger.LogDebug($"На листе {sheet.Name} найдено {sheetResults.Count} дисциплин");
                allResults.AddRange(sheetResults);
            }

            logger.LogInformation($"Всего собрано оценок до фильтрации: {allResults.Count}");

            // Разделяем дисциплины на уникальные (фильтруемые) и не уникальные (практики и курсовые)
            var uniqueDisciplines = new List<StudentDisciplineResult>();
            var nonUniqueDisciplines = new List<StudentDisciplineResult>();

            foreach (var result in allResults)
            {
                string disciplineName = result.DisciplineName ?? "";
                // Фильтр для практик и спец. имен (добавили ВКР/Гос чтобы они не схлопнулись случайно, если попадут сюда)
                if (disciplineName.Contains("НАЗВАНИЕ ДИСЦИПЛИНЫ") ||
                    disciplineName.Contains("Выпускная квалификационная работа", StringComparison.OrdinalIgnoreCase) ||
                    disciplineName.Contains("Государственный экзамен", StringComparison.OrdinalIgnoreCase) ||
                    disciplineName.Contains("Учебная практик", StringComparison.OrdinalIgnoreCase) ||
                    disciplineName.Contains("научно-исследоват", StringComparison.OrdinalIgnoreCase) ||
                    disciplineName.Contains("Производственная практик", StringComparison.OrdinalIgnoreCase) ||
                    disciplineName.Contains("Практика по получению профессиональ", StringComparison.OrdinalIgnoreCase))
                {
                    nonUniqueDisciplines.Add(result);
                }
                else
                {
                    uniqueDisciplines.Add(result);
                }
            }

            var filteredUnique = uniqueDisciplines
                .GroupBy(r => r.DisciplineName)
                .Select(g => g.OrderByDescending(r => r.Year ?? 0).ThenByDescending(r => r.Semester).First())
                .ToList();

            var finalResults = filteredUnique.Concat(nonUniqueDisciplines).ToList();
            bool withHonors = CalculateDiplomaWithHonors(finalResults);

            return (finalResults, withHonors);
        }

        private List<StudentDisciplineResult> ParseEducationResultsWorksheet(ExcelWorksheet sheet)
        {
            logger.LogDebug($"Начало парсинга листа: {sheet.Name}");
            var results = new List<StudentDisciplineResult>();

            var tableHeaderCells = sheet.Cells
                .Where(c => c.Value != null && c.Text.Contains("п/п"))
                .Reverse();

            foreach (var headerCell in tableHeaderCells)
            {
                int col = headerCell.Start.Column;
                int row = headerCell.Start.Row;

                var semesterVal = sheet.Cells[row - 2, col + 5].Value;
                short semester = 0;
                if (semesterVal != null) short.TryParse(semesterVal.ToString(), out semester);

                row += 1;
                while (sheet.Cells[row, col].Value is null) row++;

                // --- 1. Обычные дисциплины ---
                while (sheet.Cells[row, col].Value is not null &&
                       sheet.Cells[row, col + 1].Value is not null &&
                       int.TryParse(sheet.Cells[row, col].Value.ToString(), out _))
                {
                    var result = new StudentDisciplineResult();
                    result.DisciplineName = sheet.Cells[row, col + 1].Text.Trim();
                    result.CreditUnits = sheet.Cells[row, col + 2].GetValue<double?>() ?? 0;
                    result.AudHours = sheet.Cells[row, col + 3].GetValue<double?>() ?? 0;
                    result.ControlType = sheet.Cells[row, col + 6].Text.Trim();
                    result.Score = sheet.Cells[row, col + 7].Text.Trim();

                    // Парсинг даты для дисциплины (col + 10 = K)
                    // Используем false, чтобы не падать при ошибке, а ставить null
                    var date = ObjectToDateOnly(sheet.Cells[row, col + 10].Value, false);
                    result.Year = date is null ? null : (short)date.Value.Year;
                    result.Semester = semester;

                    results.Add(result);
                    row++;
                }

                // --- 2. КУРСОВЫЕ РАБОТЫ ---
                int maxSearchRow = Math.Min(row + 20, sheet.Dimension.End.Row);

                for (int i = row; i < maxSearchRow; i++)
                {
                    var textA = sheet.Cells[i, col].Text.Trim();
                    var textB = sheet.Cells[i, col + 1].Text.Trim();

                    // Ищем строку с заголовком "Курсовая работа"
                    if ((!string.IsNullOrEmpty(textA) && (textA.Contains("Курсовая работа") || textA.Contains("Курсовой проект"))) ||
                        (!string.IsNullOrEmpty(textB) && (textB.Contains("Курсовая работа") || textB.Contains("Курсовой проект"))))
                    {
                        // 1. Тема
                        string topic = "Тема курсовой работы";
                        int topicRow = i + 1;

                        var cellTopicB = sheet.Cells[topicRow, col + 1].Text.Trim();
                        var cellTopicC = sheet.Cells[topicRow, col + 2].Text.Trim();

                        if (cellTopicB.Contains("по теме", StringComparison.OrdinalIgnoreCase))
                        {
                            topic = (cellTopicB.Length < 15 && !string.IsNullOrWhiteSpace(cellTopicC)) ? cellTopicC : cellTopicB;
                        }
                        else if (!string.IsNullOrWhiteSpace(cellTopicC) && !cellTopicC.StartsWith("#") && !cellTopicC.Contains("Error"))
                        {
                            topic = cellTopicC;
                        }
                        else if (!string.IsNullOrWhiteSpace(cellTopicB) && !cellTopicB.StartsWith("#"))
                        {
                            topic = cellTopicB;
                        }

                        topic = Regex.Replace(topic, @"^по\s+теме:?\s*", "", RegexOptions.IgnoreCase).Trim().Trim('"');
                        if (string.IsNullOrWhiteSpace(topic) || topic.StartsWith("#")) topic = "Тема курсовой работы";

                        string disciplineName = $"НАЗВАНИЕ ДИСЦИПЛИНЫ \"{topic}\"";

                        // 2. Смещения колонок
                        // ControlType: col + 6 (G)
                        // Score:       col + 7 (H)
                        // Date:        col + 9 (J)

                        // Принудительно "Курсовая работа", игнорируем "экзамен" из ячейки
                        string controlType = "Курсовая работа";

                        // Получаем оценку
                        string scoreRaw = sheet.Cells[i, col + 7].Text.Trim();
                        if (string.IsNullOrWhiteSpace(scoreRaw) || scoreRaw.StartsWith("#")) scoreRaw = "0";

                        // 3. Парсинг года (col + 9)
                        var dateCellVal = sheet.Cells[i, col + 9].Value;
                        short? courseYear = null;

                        // Используем наш обновленный безопасный метод
                        var courseDateObj = ObjectToDateOnly(dateCellVal, false);
                        if (courseDateObj != null)
                        {
                            courseYear = (short)courseDateObj.Value.Year;
                        }

                        var courseResult = new StudentDisciplineResult
                        {
                            DisciplineName = disciplineName,
                            CreditUnits = 0,
                            AudHours = 0,
                            ControlType = controlType,
                            Score = scoreRaw,
                            Year = courseYear,
                            Semester = semester
                        };

                        results.Add(courseResult);
                        logger.LogDebug($"Добавлена курсовая: {disciplineName}, Оценка: {scoreRaw}, Год: {courseYear}");

                        i++;
                    }
                }
            }

            return results;
        }

        private bool CalculateDiplomaWithHonors(List<StudentDisciplineResult> disciplineResults)
        {
            logger.LogInformation("Начало расчета возможности получения диплома с отличием");
            logger.LogDebug($"Всего дисциплин для анализа: {disciplineResults.Count}");

            // 1. Проверяем НАЛИЧИЕ ПЛОХИХ РЕЗУЛЬТАТОВ в любой дисциплине
            bool hasBadResults = HasBadResults(disciplineResults);
            if (hasBadResults)
            {
                logger.LogWarning("Есть плохие результаты (незачеты/неуды) - диплом с отличием невозможен");
                return false;
            }

            // 2. Проверяем итоговые государственные аттестации
            var finalAttestations = disciplineResults
                .Where(r => IsFinalStateAttestation(r.DisciplineName))
                .ToList();

            logger.LogDebug($"Найдено итоговых государственных аттестаций: {finalAttestations.Count}");

            if (finalAttestations.Any())
            {
                logger.LogDebug("Проверка оценок итоговых аттестаций:");
                foreach (var attestation in finalAttestations)
                {
                    if (!IsExcellentOrPassGrade(attestation.Score))
                    {
                        logger.LogWarning($"Итоговая аттестация '{attestation.DisciplineName}' имеет оценку '{attestation.Score}' вместо 'отлично/зачет'");
                        return false;
                    }
                    logger.LogTrace($"Аттестация: {attestation.DisciplineName}, Оценка: {attestation.Score} - OK");
                }
                logger.LogDebug("Все итоговые аттестации сданы на 'отлично/зачет'");
            }
            else
            {
                logger.LogWarning("Не найдено итоговых государственных аттестаций");
                return false; // Без итоговых аттестаций диплом с отличием невозможен
            }

            // 3. Отбираем дисциплины для процентного подсчета
            var countedDisciplines = disciplineResults
                .Where(r => ShouldCountInPercentage(r))
                .ToList();

            logger.LogDebug($"Дисциплин для процентного подсчета: {countedDisciplines.Count}");

            if (!countedDisciplines.Any())
            {
                logger.LogWarning("Нет дисциплин для процентного подсчета");
                return false;
            }

            // Логирование деталей по каждой дисциплине
            logger.LogTrace("Детали по отобранным дисциплинам:");
            foreach (var disc in countedDisciplines)
            {
                logger.LogTrace($"  - {disc.DisciplineName}: {disc.Score} ({disc.ControlType})");
            }

            // 4. Подсчитываем оценки в процентном соотношении
            int totalCount = countedDisciplines.Count;
            int excellentCount = countedDisciplines.Count(r => IsExcellentGrade(r.Score));
            int goodCount = countedDisciplines.Count(r => IsGoodGrade(r.Score));
            int satisfactoryCount = countedDisciplines.Count(r => IsSatisfactoryGrade(r.Score));

            logger.LogDebug($"Статистика оценок для процентов:");
            logger.LogDebug($"  Всего: {totalCount}");
            logger.LogDebug($"  Отлично (13-15): {excellentCount}");
            logger.LogDebug($"  Хорошо (10-12): {goodCount}");
            logger.LogDebug($"  Удовлетворительно (7-9): {satisfactoryCount}");

            // Проверяем наличие удовлетворительных оценок
            if (satisfactoryCount > 0)
            {
                logger.LogWarning($"Найдены удовлетворительные оценки. Их не должно быть для диплома с отличием");
                return false;
            }

            // 5. Рассчитываем проценты
            double excellentPercentage = (double)excellentCount / totalCount * 100;
            double goodPercentage = (double)goodCount / totalCount * 100;

            logger.LogDebug($"Проценты: Отлично = {excellentPercentage:F2}%, Хорошо = {goodPercentage:F2}%");

            // 6. Проверяем условия
            bool meetsExcellentCriteria = excellentPercentage >= 75;
            bool meetsGoodCriteria = goodPercentage <= 25;
            bool onlyExcellentAndGood = (excellentCount + goodCount) == totalCount;

            logger.LogDebug($"Критерии:");
            logger.LogDebug($"  Не менее 75% отлично: {(meetsExcellentCriteria ? "ДА" : "НЕТ")} ({excellentPercentage:F2}%)");
            logger.LogDebug($"  Не более 25% хорошо: {(meetsGoodCriteria ? "ДА" : "НЕТ")} ({goodPercentage:F2}%)");
            logger.LogDebug($"  Только отлично и хорошо: {(onlyExcellentAndGood ? "ДА" : "НЕТ")}");

            bool result = meetsExcellentCriteria && meetsGoodCriteria && onlyExcellentAndGood;

            logger.LogInformation($"Итоговый результат расчета диплома с отличием: {(result ? "ДА" : "НЕТ")}");

            return result;
        }

        private bool HasBadResults(List<StudentDisciplineResult> disciplineResults)
        {
            logger.LogDebug("Проверка на наличие плохих результатов во всех дисциплинах");

            foreach (var result in disciplineResults)
            {
                string disciplineName = result.DisciplineName ?? "";
                string score = result.Score ?? "";
                string controlType = result.ControlType?.ToLower() ?? "";

                // 1. Проверка практик - ТОЛЬКО НЕУД
                if (disciplineName.Contains("Учебная практика", StringComparison.OrdinalIgnoreCase) ||
                    disciplineName.Contains("Производственная практика", StringComparison.OrdinalIgnoreCase))
                {
                    // Для практик: числовая оценка 0-6 или "неуд"
                    if (IsUnsatisfactoryGrade(score))
                    {
                        logger.LogWarning($"Практика '{disciplineName}' имеет неудовлетворительную оценку: {score}");
                        return true;
                    }
                    continue;
                }

                // 2. Проверка недифференцированных зачетов
                if (controlType.Contains("зачет"))
                {
                    // Проверяем, является ли зачет недифференцированным
                    if (!IsDifferentiatedCredit(score))
                    {
                        // Недифференцированный зачет - только "незачет" плохо
                        if (IsFailGrade(score))
                        {
                            logger.LogWarning($"Недифференцированный зачет '{disciplineName}' имеет оценку 'незачет': {score}");
                            return true;
                        }
                    }
                    else
                    {
                        // Дифференцированный зачет - проверяем на неуд (0-6 баллов)
                        if (IsUnsatisfactoryGrade(score))
                        {
                            logger.LogWarning($"Дифференцированный зачет '{disciplineName}' имеет неудовлетворительную оценку: {score}");
                            return true;
                        }
                    }
                    continue;
                }

                // 3. Проверка ВСЕХ остальных дисциплин на неудовлетворительные оценки
                if (IsUnsatisfactoryGrade(score))
                {
                    logger.LogWarning($"Дисциплина '{disciplineName}' имеет неудовлетворительную оценку: {score}");
                    return true;
                }
            }

            logger.LogDebug("Плохих результатов не найдено");
            return false;
        }

        // Вспомогательные методы с логированием
        private bool IsFinalStateAttestation(string disciplineName)
        {
            bool result = disciplineName.ToLower().Contains("итоговый государственный экзамен") ||
                          disciplineName.ToLower().Contains("государственный экзамен") ||
                          disciplineName.ToLower().Contains("защита выпускной квалификационной работы") ||
                          disciplineName.ToLower().Contains("вкр") ||
                          disciplineName.ToLower().Contains("выпускная квалификационная работа");

            if (result)
                logger.LogTrace($"Дисциплина '{disciplineName}' определена как итоговая аттестация");

            return result;
        }

        private bool ShouldCountInPercentage(StudentDisciplineResult result)
        {
            string disciplineName = result.DisciplineName ?? "";
            string controlType = result.ControlType?.ToLower() ?? "";
            string score = result.Score ?? "";

            // 1. Практики не учитываются в процентах
            if (disciplineName.Contains("Учебная практика", StringComparison.OrdinalIgnoreCase) ||
                disciplineName.Contains("Производственная практика", StringComparison.OrdinalIgnoreCase))
            {
                logger.LogTrace($"Дисциплина '{disciplineName}' - практика, не учитывается в процентах");
                return false;
            }

            // 2. Итоговые аттестации проверяются отдельно
            if (IsFinalStateAttestation(disciplineName))
            {
                logger.LogTrace($"Дисциплина '{disciplineName}' - итоговая аттестация, не учитывается в процентах");
                return false;
            }

            // 3. Зачеты
            if (controlType.Contains("зачет"))
            {
                // Если есть числовая оценка 0-15 - это дифференцированный зачет
                if (int.TryParse(score, out int numericScore))
                {
                    // Числовой зачет 0-15 - дифференцированный, учитывается в процентах
                    logger.LogTrace($"Дисциплина '{disciplineName}' - дифференцированный зачет (числовая оценка: {numericScore}), учитывается");
                    return true;
                }

                // Если словесная оценка
                var lowerScore = score.ToLower().Trim();

                // "Зачет"/"незачет" - недифференцированный зачет, не учитывается
                if (lowerScore == "зачет" || lowerScore == "незачет")
                {
                    logger.LogTrace($"Дисциплина '{disciplineName}' - недифференцированный зачет, не учитывается");
                    return false;
                }

                // Любая другая словесная оценка ("отлично", "хорошо", "удовл") - дифференцированный зачет
                if (!string.IsNullOrWhiteSpace(score))
                {
                    logger.LogTrace($"Дисциплина '{disciplineName}' - дифференцированный зачет (словесная оценка: {score}), учитывается");
                    return true;
                }

                // Пустая оценка для зачета - недифференцированный, не учитывается
                logger.LogTrace($"Дисциплина '{disciplineName}' - зачет без оценки, не учитывается");
                return false;
            }

            // 4. Экзамены всегда учитываются (включая курсовые работы)
            if (controlType.Contains("экзамен"))
            {
                logger.LogTrace($"Дисциплина '{disciplineName}' - экзамен, учитывается");
                return true;
            }

            // 5. Курсовые работы с оценкой (если не помечены как экзамен)
            if (controlType.Contains("курсов"))
            {
                if (!string.IsNullOrWhiteSpace(score))
                {
                    logger.LogTrace($"Дисциплина '{disciplineName}' - курсовая работа с оценкой, учитывается");
                    return true;
                }
                else
                {
                    logger.LogTrace($"Дисциплина '{disciplineName}' - курсовая работа без оценки, не учитывается");
                    return false;
                }
            }

            logger.LogTrace($"Дисциплина '{disciplineName}' - форма контроля '{controlType}' не учитывается в процентах");
            return false;
        }

        private bool IsDifferentiatedCredit(string score)
        {
            if (string.IsNullOrWhiteSpace(score))
                return false;

            // Дифференцированный зачет имеет числовую оценку 0-15
            if (int.TryParse(score, out int numericScore))
            {
                return numericScore >= 0 && numericScore <= 15;
            }

            // ИЛИ словесную оценку "отлично", "хорошо", "удовл"
            var lowerScore = score.ToLower().Trim();
            return lowerScore == "отлично" || lowerScore == "отл" ||
                   lowerScore == "хорошо" || lowerScore == "хор" ||
                   lowerScore == "удовлетворительно" || lowerScore == "удовл";
        }

        private bool IsFailGrade(string score)
        {
            if (string.IsNullOrWhiteSpace(score))
                return false;

            var lowerScore = score.ToLower().Trim();

            // "Незачет" - плохо для недифференцированного зачета
            if (lowerScore == "незачет")
                return true;

            // Числовой 0 в недифференцированном зачете - незачет
            // Но мы проверяем это только если IsDifferentiatedCredit вернул false
            if (int.TryParse(score, out int numericScore))
            {
                return numericScore == 0; // 0 = незачет
            }

            return false;
        }

        private bool IsUnsatisfactoryGrade(string score)
        {
            if (string.IsNullOrWhiteSpace(score))
                return false;

            var lowerScore = score.ToLower().Trim();

            // Словесные оценки
            if (lowerScore == "неудовлетворительно" || lowerScore == "неуд")
            {
                return true;
            }

            // Числовые оценки в 15-балльной системе
            if (int.TryParse(score, out int numericScore))
            {
                // 0-6 баллов = неудовлетворительно
                return numericScore >= 0 && numericScore <= 6;
            }

            return false;
        }

        private bool IsExcellentOrPassGrade(string score)
        {
            // Проверка на отлично (13-15, "отлично") ИЛИ зачет (15, "зачтено")
            if (IsExcellentGrade(score))
                return true;

            var lowerScore = score.ToLower().Trim();
            if (lowerScore == "зачтено" || lowerScore == "зачет")
                return true;

            if (int.TryParse(score, out int numericScore))
            {
                return numericScore == 15; // 15 баллов = зачет
            }

            return false;
        }

        private bool IsExcellentGrade(string score)
        {
            if (string.IsNullOrWhiteSpace(score))
                return false;

            // Если оценка в виде слова
            var lowerScore = score.ToLower().Trim();
            if (lowerScore == "отлично" || lowerScore == "отл")
            {
                return true;
            }

            // Если оценка в виде числа
            if (int.TryParse(score, out int numericScore))
            {
                // Для обычных дисциплин: 13-15 баллов
                // Для защиты ВКР: 5 баллов (если 5-балльная система)
                return numericScore >= 13 && numericScore <= 15 || numericScore == 5;
            }

            return false;
        }

        private bool IsGoodGrade(string score)
        {
            if (string.IsNullOrWhiteSpace(score))
                return false;

            var lowerScore = score.ToLower().Trim();
            if (lowerScore == "хорошо" || lowerScore == "хор" || lowerScore == "4")
            {
                return true;
            }

            if (int.TryParse(score, out int numericScore))
            {
                // Для обычных дисциплин: 10-12 баллов
                // Для защиты ВКР: 4 балла (если 5-балльная система)
                return numericScore >= 10 && numericScore <= 12 || numericScore == 4;
            }

            return false;
        }

        private bool IsSatisfactoryGrade(string score)
        {
            if (string.IsNullOrWhiteSpace(score))
            {
                logger.LogTrace($"Пустая оценка, не является 'удовлетворительно'");
                return false;
            }

            var lowerScore = score.ToLower().Trim();
            if (lowerScore == "удовлетворительно" || lowerScore == "удовл" || lowerScore == "3")
            {
                logger.LogTrace($"Оценка '{score}' определена как 'удовлетворительно' (словесная)");
                return true;
            }

            if (int.TryParse(score, out int numericScore))
            {
                bool result = numericScore >= 7 && numericScore <= 9;
                if (result)
                    logger.LogTrace($"Оценка '{score}' ({numericScore}) определена как 'удовлетворительно' (7-9 баллов)");
                return result;
            }

            return false;
        }

        /// <summary>
        /// Маппинг словесных оценок из карточки ("5, отлично") в 15-балльную шкалу.
        /// </summary>
        private string MapGradeTo15Scale(string rawScore)
        {
            if (string.IsNullOrWhiteSpace(rawScore)) return "";
            var lower = rawScore.ToLower().Trim();

            // 1. Проверяем наличие цифр 5, 4, 3, 2
            if (lower.Contains("5")) return "15";
            if (lower.Contains("4")) return "12";
            if (lower.Contains("3")) return "9";
            if (lower.Contains("2")) return "2";

            // 2. Если цифр нет, проверяем слова
            // "зачтено" (без цифр) считаем как 5 (15) для итоговых аттестаций (обычно они дифференцированы, но если нет - max балл)
            if (lower.Contains("зачтено") && !lower.Contains("не")) return "15";
            if (lower.Contains("отлично")) return "15";
            if (lower.Contains("хорошо")) return "12";
            if (lower.Contains("удовлетворительно")) return "9";

            // "не зачтено" или "не удовлетворительно"
            if (lower.Contains("не")) return "2";

            return rawScore; // Возвращаем как есть, если не распознали
        }

        /// <summary>Дата для приложения диплома: сначала Value ячейки, если пусто — Text.</summary>
        private DateOnly? ReadSupplementSheetDateCell(ExcelWorksheet sheet, int row, int col)
        {
            var cell = sheet.Cells[row, col];
            var fromValue = ObjectToDateOnly(cell.Value, throwEx: false);
            if (fromValue.HasValue)
                return fromValue;
            var t = cell.Text?.Trim();
            if (!string.IsNullOrWhiteSpace(t))
                return ObjectToDateOnly(t, throwEx: false);
            return null;
        }

        /// <summary>
        /// Метод, пытающийся вытащить дату из произвольного объекта или выдает ошибку.
        /// Примечание: метод может не использоваться во всех местах ImportService - где-то может 
        /// остаться хвост из такой же логики, без использования этого централизированного метода.
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="throwEx">Выкидывать ли исключение, если не удалось распарсить</param>
        /// <returns>DateOnly</returns>
        /// <exception cref="FormatException"></exception>
        private DateOnly? ObjectToDateOnly(object obj, bool throwEx = true)
        {
            if (obj is null || string.IsNullOrWhiteSpace(obj.ToString()))
                return throwEx ? throw new FormatException("Дата пуста") : null;

            // 1. Если Excel вернул DateTime
            if (obj is DateTime date)
                return DateOnly.FromDateTime(date);

            string dateStr = obj.ToString().Trim();

            // 2. Если Excel вернул число (OLE Automation Date)
            if (double.TryParse(dateStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double oleDate))
            {
                return DateOnly.FromDateTime(DateTime.FromOADate(oleDate));
            }

            // 3. Парсинг строки по конкретным шаблонам (Invariant Mode Safe)
            // Учитываем варианты с ведущими нулями и без (д.М.гггг)
            string[] formats = {
                "dd.MM.yyyy", "d.M.yyyy",
                "d.MM.yyyy",  "dd.M.yyyy",
                "yyyy-MM-dd" // на всякий случай ISO
            };

            if (DateTime.TryParseExact(dateStr, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate))
            {
                return DateOnly.FromDateTime(parsedDate);
            }

            return throwEx ? throw new FormatException($"Не удалось распознать дату: {obj}") : null;
        }
    }
}
