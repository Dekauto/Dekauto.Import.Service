using Dekauto.Import.Service.Domain.Entities;
using Dekauto.Import.Service.Domain.Interfaces;
using OfficeOpenXml;
using System.Diagnostics.Contracts;
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
                                    case "№ договора": case "номер договора":
                                        if (isCurrentStudent == true)
                                        {
                                            student.EducationRelationNum = cellValue.ToString();
                                        }
                                        break;
                                    case "№ приказа о зачислении": case "номер приказа о зачислении":
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
                            var header = headers[col-1];
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
                                    string name = $"{fio[1].ToString().Substring(0,1).ToUpper()}{fio[1].ToString().Substring(1)}";
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
                                    if (cellValue.ToString() != "") { 
                                        student.AddressResidentialIndex = Regex.Match(cellValue.ToString(), indexPattern).ToString();
                                        student.AddressResidentialCity = Regex.Match(cellValue.ToString(), cityPattern).Groups[1].ToString();
                                        switch(Regex.Match(cellValue.ToString(), addressTypePattern).Groups[1].ToString()) 
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
                                    else { 
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
                            if ((header.ToLower() == "адрес проживания" ) && (cellValue.ToString() == "")) 
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
                        collapsed = Regex.Replace(collapsed, @"\(\s*Итоговая\s+контрольная\s+работа\s*\)", string.Empty, RegexOptions.IgnoreCase);
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

                            // Поиск номера курса в 4-й строке по всем столбцам
                            int? courseNum = null;
                            for (int col = 1; col <= columnCount; col++)
                            {
                                var courseCellText = GetMergedText(4, col);
                                if (TryExtractFirstInt(courseCellText, out var num) && num >= 1 && num <= 6)
                                {
                                    courseNum = num;
                                    break;
                                }
                            }

                            if (courseNum.HasValue && courseNum.Value > 0)
                            {
                                var sem = courseNum.Value * 2;
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
    }
}
