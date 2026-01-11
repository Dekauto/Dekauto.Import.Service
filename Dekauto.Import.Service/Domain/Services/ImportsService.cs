using Dekauto.Import.Service.Domain.Entities;
using Dekauto.Import.Service.Domain.Entities.DTO;
using Dekauto.Import.Service.Domain.Interfaces;
using OfficeOpenXml;
using System.Globalization;
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
        public async Task<IEnumerable<Student>> GetStudentsContract(IFormFile contract, List<Student> students)
        {
            using (var stream = new MemoryStream())
            {
                await contract.CopyToAsync(stream);
                using (var packege = new ExcelPackage(stream))
                {
                    var worksheet = packege.Workbook.Worksheets[0] ?? throw new InvalidOperationException("Загруженный файл не содержит листов");

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

                using (var packege = new ExcelPackage(stream))
                {
                    if (packege.Workbook.Worksheets.Count == 0)
                        throw new InvalidOperationException("Загруженный файл не содержит листов");

                    var worksheet = packege.Workbook.Worksheets["ПланСвод"]
                        ?? throw new InvalidOperationException("Загруженный файл не содержит листа ПланСвод");

                    if (worksheet.Dimension == null)
                        return students;

                    var columnCount = worksheet.Dimension.Columns;
                    var rowCount = worksheet.Dimension.Rows;

                    string GetMergedText(int row, int col)
                    {
                        var mergedAddress = worksheet.MergedCells[row, col];
                        if (!string.IsNullOrWhiteSpace(mergedAddress))
                            return worksheet.Cells[mergedAddress].First().Text;
                        return worksheet.Cells[row, col].Text;
                    }

                    static string NormalizePlanHeader(string? value)
                    {
                        if (string.IsNullOrWhiteSpace(value))
                            return string.Empty;
                        return Regex.Replace(value, @"\s+", " ").Trim();
                    }

                    static bool ContainsExpertHeader(string value)
                    {
                        if (string.IsNullOrWhiteSpace(value))
                            return false;
                        var normalized = value.ToLower().Replace(" ", string.Empty);
                        return normalized.Contains("экспертное");
                    }

                    int? audHoursCol = null;
                    var creditUnitsBySemesterCol = new Dictionary<int, int>();

                    for (int col = 1; col <= columnCount; col++)
                    {
                        var header2 = NormalizePlanHeader(GetMergedText(2, col));
                        var header3 = NormalizePlanHeader(GetMergedText(3, col));

                        var header2Lower = header2.ToLower();
                        var header3Lower = header3.ToLower();

                        // "Ауд." находится в 3 строке заголовков, под ним — нужные значения для дисциплин
                        // Обычно это колонка в блоке "Итого ... часов"
                        if (audHoursCol == null &&
                            (header2Lower.Contains("итого") || header2Lower.Contains("всего")) &&
                            Regex.IsMatch(header3Lower, @"\bауд\.?\b", RegexOptions.IgnoreCase))
                        {
                            audHoursCol = col;
                            continue;
                        }

                        var matchSemester = Regex.Match(header2Lower, @"семестр\s*(\d{1,2})");
                        if (matchSemester.Success && (header3Lower.Contains("з.е") || header3Lower.Contains("з. е")))
                        {
                            if (int.TryParse(matchSemester.Groups[1].Value, out var sem) && sem >= 1 && sem <= 8)
                            {
                                creditUnitsBySemesterCol[sem] = col;
                            }
                        }
                    }

                    if (audHoursCol == null || creditUnitsBySemesterCol.Count == 0)
                        throw new InvalidOperationException("Не удалось определить колонки учебного плана (Ауд. часов/з.е. по семестрам)");

                    static string NormalizeDisciplineName(string? value)
                    {
                        if (string.IsNullOrWhiteSpace(value))
                            return string.Empty;
                        return Regex.Replace(value, @"\s+", " ").Trim();
                    }

                    bool TryGetIntCell(int row, int col, out int result)
                    {
                        result = default;
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
                        if (double.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed))
                        {
                            result = (int)Math.Round(parsed);
                            return true;
                        }
                        return false;
                    }

                    bool TryGetDoubleCell(int row, int col, out double result)
                    {
                        result = default;
                        var value = worksheet.Cells[row, col].Value;
                        if (value == null)
                            return false;

                        if (value is double d)
                        {
                            result = d;
                            return true;
                        }
                        if (value is float f)
                        {
                            result = f;
                            return true;
                        }
                        if (value is decimal dec)
                        {
                            result = (double)dec;
                            return true;
                        }
                        if (value is int i)
                        {
                            result = i;
                            return true;
                        }
                        if (value is long l)
                        {
                            result = l;
                            return true;
                        }

                        var str = value.ToString()?.Trim();
                        if (string.IsNullOrWhiteSpace(str))
                            return false;
                        str = str.Replace(" ", string.Empty).Replace(",", ".");
                        return double.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out result);
                    }

                    // Строки с дисциплинами начинаются с 6 строки, название — в 4 столбце
                    for (int row = 6; row <= rowCount; row++)
                    {
                        var nameRaw = worksheet.Cells[row, 4].Text;
                        var disciplineName = NormalizeDisciplineName(nameRaw);
                        if (string.IsNullOrWhiteSpace(disciplineName))
                            continue;

                        var isModuleRow = worksheet.Cells[row, 4].Style.Font.Bold;
                        if (isModuleRow)
                            continue;

                        int? audHours = null;
                        if (TryGetIntCell(row, audHoursCol.Value, out var hours))
                            audHours = hours;

                        foreach (var kvp in creditUnitsBySemesterCol)
                        {
                            var sem = (short)kvp.Key;
                            var col = kvp.Value;
                            if (!TryGetDoubleCell(row, col, out var ze))
                                continue;

                            foreach (var student in students)
                            {
                                var target = student.DisciplineResults.FirstOrDefault(x =>
                                    x.DisciplineName != null &&
                                    x.Semester.HasValue &&
                                    x.Semester.Value == sem &&
                                    x.DisciplineName.Equals(disciplineName, StringComparison.OrdinalIgnoreCase));

                                if (target == null)
                                    continue;

                                target.AudHours = audHours;
                                target.CreditUnits = ze;
                            }
                        }
                    }

                    // Извлечение типа контроля из листов Курс 1, Курс 2, Курс 3, Курс 4
                    // Столбец 5 - название дисциплины
                    // Столбец 7 - вид контроля для нечётных семестров (1, 3, 5, 7)
                    // Столбец 37 - вид контроля для чётных семестров (2, 4, 6, 8)
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
                            "к" => "контрольная",
                            "кр" => "курсовая",
                            "за" => "зачёт",
                            "зао" => "зачёт с оценкой",
                            _ => value.Trim()
                        };
                    }

                    var courseSheets = new[] { "Курс 1", "Курс 2", "Курс 3", "Курс 4" };
                    for (int courseIndex = 0; courseIndex < courseSheets.Length; courseIndex++)
                    {
                        var courseSheet = packege.Workbook.Worksheets[courseSheets[courseIndex]];
                        if (courseSheet?.Dimension == null)
                            continue;

                        var courseNum = courseIndex + 1;
                        var oddSemester = (short)(courseNum * 2 - 1);  // 1, 3, 5, 7
                        var evenSemester = (short)(courseNum * 2);     // 2, 4, 6, 8

                        var courseRowCount = courseSheet.Dimension.Rows;

                        for (int row = 1; row <= courseRowCount; row++)
                        {
                            var nameRaw = courseSheet.Cells[row, 5].Text;
                            var disciplineName = NormalizeDisciplineName(nameRaw);
                            if (string.IsNullOrWhiteSpace(disciplineName))
                                continue;

                            // Вид контроля для нечётного семестра (столбец 7)
                            var oddControlRaw = courseSheet.Cells[row, 7].Text;
                            var oddControlType = NormalizeControlType(oddControlRaw);

                            // Вид контроля для чётного семестра (столбец 22)
                            var evenControlRaw = courseSheet.Cells[row, 22].Text;
                            var evenControlType = NormalizeControlType(evenControlRaw);

                            foreach (var student in students)
                            {
                                // Обновление для нечётного семестра
                                if (!string.IsNullOrWhiteSpace(oddControlType))
                                {
                                    var targetOdd = student.DisciplineResults.FirstOrDefault(x =>
                                        x.DisciplineName != null &&
                                        x.Semester.HasValue &&
                                        x.Semester.Value == oddSemester &&
                                        x.DisciplineName.Equals(disciplineName, StringComparison.OrdinalIgnoreCase));

                                    if (targetOdd != null && string.IsNullOrWhiteSpace(targetOdd.ControlType))
                                    {
                                        targetOdd.ControlType = oddControlType;
                                    }
                                }

                                // Обновление для чётного семестра
                                if (!string.IsNullOrWhiteSpace(evenControlType))
                                {
                                    var targetEven = student.DisciplineResults.FirstOrDefault(x =>
                                        x.DisciplineName != null &&
                                        x.Semester.HasValue &&
                                        x.Semester.Value == evenSemester &&
                                        x.DisciplineName.Equals(disciplineName, StringComparison.OrdinalIgnoreCase));

                                    if (targetEven != null && string.IsNullOrWhiteSpace(targetEven.ControlType))
                                    {
                                        targetEven.ControlType = evenControlType;
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return students;
        }

        public async Task<IEnumerable<Student>> GetStudentsJournal(IFormFile journal, List<Student> students)
        {
            using (var stream = new MemoryStream())
            {
                await journal.CopyToAsync(stream);
                using (var packege = new ExcelPackage(stream))
                {
                    var worksheet = packege.Workbook.Worksheets[0] ?? throw new InvalidOperationException("Загруженный файл не содержит листов");

                    var columnCount = worksheet.Dimension.Columns;
                    var rowCount = worksheet.Dimension.Rows;

                    var headers = new List<string>();
                    for (int col = 1; col <= columnCount; col++)
                    {
                        headers.Add(worksheet.Cells[1, col].Text);
                    }

                    foreach (var student in students)
                    {
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
                    }
                }
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
                            string addressTypePattern = @"\b(?:\w+\s+(?<abbr>[гсхдп])\b|(?<abbr>[гсхдп])\.?\s+\w+\b)";
                            string cityPattern = @"\b(?:(?<city>[\w\s]+?)\s+[гсхдп]\,?|(?<type>[гсхдп])\.?\s*(?<city>[\w\s]+?))\b";
                            string housePattern = @"(?:\bдом|д)\.?\s*(\w+)\b,?|\bд(\w+)\b";
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
                                        student.AddressRegistrationIndex = Regex.Match(cellValue.ToString(), indexPattern).ToString();
                                        student.AddressRegistrationCity = Regex.Match(cellValue.ToString(), cityPattern).Groups[1].ToString();
                                        switch (Regex.Match(cellValue.ToString(), addressTypePattern).Groups[1].ToString())
                                        {
                                            case "г":
                                                student.AddressRegistrationType = "город";
                                                break;
                                            case "с":
                                                student.AddressRegistrationType = "село";
                                                break;
                                            case "х":
                                                student.AddressRegistrationType = "хутор";
                                                break;
                                            case "д":
                                                student.AddressRegistrationType = "деревня";
                                                break;
                                            case "п":
                                                student.AddressRegistrationType = "посёлок";
                                                break;
                                        }
                                        student.AddressRegistrationHouse = Regex.Match(cellValue.ToString(), housePattern).Groups[1].ToString();
                                        student.AddressRegistrationStreet = Regex.Match(cellValue.ToString(), streetPattern).Groups[1].ToString().Trim();
                                        string housingMatch = Regex.Match(cellValue.ToString(), housingTypePattern).Groups[1].ToString().Trim();
                                        if (housingMatch == "к" || housingMatch == "к." || housingMatch == "корпус" || housingMatch == "корпус.")
                                            student.AddressRegistrationHousingType = "корпус";
                                        else
                                            if (housingMatch == "стр" || housingMatch == "стр." || housingMatch == "строение" || housingMatch == "строение.")
                                            student.AddressRegistrationHousingType = "строение";
                                        student.AddressRegistrationHousing = Regex.Match(cellValue.ToString(), housingPattern).Groups[2].ToString().Trim();
                                        student.AddressRegistrationApartment = Regex.Match(cellValue.ToString(), apartementPattern).Groups[1].ToString().Trim();
                                        student.AddressRegistrationOblKrayAvtobl = Regex.Match(cellValue.ToString(), regionPattern).Groups[1].ToString().Trim();

                                    }
                                    break;

                                case "адрес проживания":
                                    if (cellValue.ToString() != "")
                                    {
                                        student.AddressResidentialIndex = Regex.Match(cellValue.ToString(), indexPattern).ToString();
                                        student.AddressResidentialCity = Regex.Match(cellValue.ToString(), cityPattern).Groups[1].ToString();
                                        switch (Regex.Match(cellValue.ToString(), addressTypePattern).Groups[1].ToString())
                                        {
                                            case "г":
                                                student.AddressResidentialType = "город";
                                                break;
                                            case "с":
                                                student.AddressResidentialType = "село";
                                                break;
                                            case "х":
                                                student.AddressResidentialType = "хутор";
                                                break;
                                            case "д":
                                                student.AddressResidentialType = "деревня";
                                                break;
                                            case "п":
                                                student.AddressResidentialType = "посёлок";
                                                break;
                                        }
                                        student.AddressResidentialHouse = Regex.Match(cellValue.ToString(), housePattern).Groups[1].ToString();
                                        student.AddressResidentialStreet = Regex.Match(cellValue.ToString(), streetPattern).Groups[1].ToString().Trim();
                                        string housingMatch = Regex.Match(cellValue.ToString(), housingTypePattern).Groups[1].ToString().Trim();
                                        if (housingMatch == "к" || housingMatch == "к." || housingMatch == "корпус" || housingMatch == "корпус.")
                                            student.AddressResidentialHousingType = "корпус";
                                        else
                                            if (housingMatch == "стр" || housingMatch == "стр." || housingMatch == "строение" || housingMatch == "строение.")
                                            student.AddressResidentialHousingType = "строение";
                                        student.AddressResidentialHousing = Regex.Match(cellValue.ToString(), housingPattern).Groups[2].ToString().Trim();
                                        student.AddressResidentialApartment = Regex.Match(cellValue.ToString(), apartementPattern).Groups[1].ToString().Trim();
                                        student.AddressResidentialOblKrayAvtobl = Regex.Match(cellValue.ToString(), regionPattern).Groups[1].ToString().Trim();

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
                using (var packege = new ExcelPackage(stream))
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
                        return Regex.Replace(collapsed, @"\s+", " ").Trim();
                    }

                    // Проверка файла на наличие листов
                    if (packege.Workbook.Worksheets.Count == 0)
                        throw new InvalidOperationException("Загруженный файл не содержит листов");

                    foreach (var worksheet in packege.Workbook.Worksheets)
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

                            var courseCellText = worksheet.Cells[4, 3].Value?.ToString();
                            if (TryExtractFirstInt(courseCellText, out var courseNum) && courseNum > 0)
                            {
                                var sem = courseNum * 2;
                                if (isAutumnWinter)
                                    sem -= 1;

                                // Если сессия не распознана, всё равно можно определить семестр по формуле для весенне-летней,
                                // но лучше оставлять null, чтобы не подставлять потенциально неверные данные.
                                if (isAutumnWinter || isSpringSummer)
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
                                var practiceMatch = Regex.Match(disciplineName, @"^(Учебная практика|Производственная практика),\s*", RegexOptions.IgnoreCase);
                                if (practiceMatch.Success)
                                {
                                    isPractice = true;
                                    // Убираем префикс "Учебная практика, " или "Производственная практика, "
                                    disciplineName = disciplineName.Substring(practiceMatch.Length).Trim();
                                    // Делаем первую букву заглавной
                                    if (!string.IsNullOrEmpty(disciplineName))
                                        disciplineName = char.ToUpper(disciplineName[0]) + disciplineName.Substring(1);
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
                                            x.DisciplineName.Equals(disciplineName, StringComparison.OrdinalIgnoreCase) &&
                                            x.ControlType != null &&
                                            x.ControlType.Equals("практика", StringComparison.OrdinalIgnoreCase))
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

        public async Task<DiplomaSupplementData> GetStudentCardAsync(IFormFile studentCard)
        {
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

                    diplomaData.BirthdayDate = ObjectToDateOnly(firstSheet.Cells[6, 5].Value);
                    diplomaData.EducationReceivedDate = ObjectToDateOnly(firstSheet.Cells[42, 8].Value);

                    // 2. Парсинг набора оценок (семестры)
                    var (disciplineGrades, honors) = ParseDisciplineGradesFromCard(package.Workbook.Worksheets);

                    // 3. Парсинг финализации (Гос. экз и ВКР) с первого листа
                    var finalizationResults = ParseFinalizationData(firstSheet);
                    if (finalizationResults.Any())
                    {
                        disciplineGrades.AddRange(finalizationResults);
                        logger.LogInformation($"Добавлено {finalizationResults.Count} записей итоговой аттестации (Гос.экзамен/ВКР).");
                    }

                    diplomaData.DisciplineResults = disciplineGrades;
                    diplomaData.DiplomaWithHonors = honors;
                }
            }

            return diplomaData;
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
                    disciplineName.Contains("Производственная практик", StringComparison.OrdinalIgnoreCase))
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

            // Отбираем ячейки начала таблиц (п/п)
            var tableHeaderCells = sheet.Cells
                .Where(c => c.Value != null && c.Text.Contains("п/п"))
                .Reverse();

            foreach (var headerCell in tableHeaderCells)
            {
                int col = headerCell.Start.Column;
                int row = headerCell.Start.Row;

                // Парсинг семестра (2 строки выше, +5 колонок)
                var semesterVal = sheet.Cells[row - 2, col + 5].Value;
                short semester = 0;
                if (semesterVal != null) short.TryParse(semesterVal.ToString(), out semester);

                row += 1; // Переход к данным

                // Пропуск пустых строк до начала списка
                while (sheet.Cells[row, col].Value is null) row++;

                // Парсинг строк таблицы
                while (sheet.Cells[row, col].Value is not null &&
                       sheet.Cells[row, col + 1].Value is not null &&
                       int.TryParse(sheet.Cells[row, col].Value.ToString(), out _))
                {
                    var result = new StudentDisciplineResult();
                    result.DisciplineName = sheet.Cells[row, col + 1].Text.Trim(); // Используем Text
                    result.CreditUnits = sheet.Cells[row, col + 2].GetValue<double?>() ?? 0;
                    result.AudHours = sheet.Cells[row, col + 3].GetValue<double?>() ?? 0;
                    result.ControlType = sheet.Cells[row, col + 6].Text.Trim();
                    result.Score = sheet.Cells[row, col + 7].Text.Trim(); // Оценка может быть строкой

                    var date = ObjectToDateOnly(sheet.Cells[row, col + 10].Value, false);
                    result.Year = date is null ? null : (short)date.Value.Year;
                    result.Semester = semester;

                    results.Add(result);
                    row++;
                }

                // --- ПОИСК КУРСОВОЙ РАБОТЫ ---
                // Ищем в диапазоне 20 строк вниз
                bool foundCourseWork = false;
                int maxSearchRow = Math.Min(row + 20, sheet.Dimension.End.Row);

                for (int i = row; i < maxSearchRow; i++)
                {
                    // Используем Text для проверки на ошибки ссылок или значения
                    var cellText = sheet.Cells[i, col].Text.Trim();

                    if (!string.IsNullOrEmpty(cellText) &&
                        (cellText.Contains("Курсовая работа") || cellText.Contains("Курсовой проект")))
                    {
                        foundCourseWork = true;

                        // 1. Извлекаем тему
                        // Тема находится на строку ниже (i+1), в следующей колонке (col+1)
                        // Если там #ССЫЛКА! или пусто -> "Тема курсовой работы"
                        var topicCell = sheet.Cells[i + 1, col + 1];
                        string topic = topicCell.Text.Trim();

                        // Проверка на ошибки Excel (#REF!, #NAME?, #ССЫЛКА! и т.д.) или пустоту
                        if (string.IsNullOrWhiteSpace(topic) || topic.StartsWith("#") || topic.Contains("Error"))
                        {
                            topic = "Тема курсовой работы";
                            logger.LogWarning($"Тема курсовой (сем. {semester}) не распознана (ошибка формулы или пусто). Установлена заглушка: {topic}");
                        }

                        // 2. Формируем название дисциплины по шаблону
                        string disciplineName = $"НАЗВАНИЕ ДИСЦИПЛИНЫ \"{topic}\"";

                        // 3. Извлекаем оценку и контроль
                        // Контроль: col + 5, Оценка: col + 6
                        string controlType = sheet.Cells[i, col + 5].Text.Trim();
                        if (string.IsNullOrEmpty(controlType) || controlType.StartsWith("#")) controlType = "экзамен"; // Фолбек

                        string score = sheet.Cells[i, col + 6].Text.Trim();
                        if (score.StartsWith("#"))
                        {
                            score = "х"; // Если оценка сломана, ставим "х" или пусто
                        }

                        // 4. Дата и год
                        var courseDate = ObjectToDateOnly(sheet.Cells[i, col + 9].Value, false);
                        short? courseYear = courseDate is null ? null : (short)courseDate.Value.Year;

                        // Добавляем результат
                        var courseResult = new StudentDisciplineResult
                        {
                            DisciplineName = disciplineName,
                            CreditUnits = 0,
                            AudHours = 0,
                            ControlType = controlType,
                            Score = score,
                            Year = courseYear,
                            Semester = semester
                        };

                        results.Add(courseResult);
                        logger.LogDebug($"Добавлена курсовая: {disciplineName}, Оценка: {score}");

                        // Прерываем поиск для этой таблицы (обычно одна курсовая на семестр в этом блоке)
                        break;
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
            var ex = new FormatException($"Не удалось распознать дату: {obj}");
            if (obj is null)
                return throwEx ? throw ex : null;

            if (obj is DateTime date)
                return DateOnly.FromDateTime(date);

            string dateStr = obj.ToString().Trim();
            // Пытаемся распарсить стандартный формат
            if (DateTime.TryParseExact(dateStr, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate))
            {
                return DateOnly.FromDateTime(parsedDate);
            }

            // Если дата пришла числом (Excel OLE Automation date)
            if (double.TryParse(dateStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double oleDate))
            {
                return DateOnly.FromDateTime(DateTime.FromOADate(oleDate));
            }

            return throwEx ? throw ex : null;
        }
    }
}
