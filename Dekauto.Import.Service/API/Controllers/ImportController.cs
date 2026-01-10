using Dekauto.Import.Service.Domain.Entities;
using Dekauto.Import.Service.Domain.Entities.Adapters;
using Dekauto.Import.Service.Domain.Interfaces;
using Dekauto.Import.Service.Domain.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace Dekauto.Import.Service.API.Controllers
{
    /// <summary>
    /// Контроллер для импорта данных студентов из файлов Excel
    /// </summary>
    /// <remarks>
    /// Этот контроллер обрабатывает загрузку и импорт данных о студентах из различных Excel файлов.
    /// Требуется авторизация через Basic Authentication.
    /// </remarks>
    [Route("api/imports")]
    [ApiController]
    [Authorize]
    public class ImportController : ControllerBase
    {
        private readonly IImportService _importService;
        private readonly ILogger<ImportController> logger;

        public ImportController(IImportService importService, ILogger<ImportController> logger)
        {
            _importService = importService ?? throw new ArgumentNullException(nameof(importService));
            this.logger = logger;
        }

        /// <summary>
        /// Импорт данных студентов из набора Excel файлов
        /// </summary>
        /// <param name="files">Адаптер с файлами для импорта</param>
        /// <returns>Список студентов с заполненными данными</returns>
        /// <remarks>
        /// Этот endpoint выполняет комплексный импорт данных о студентах из нескольких файлов Excel.
        /// 
        /// **Необходимые файлы:**
        /// - **ld** (Личное дело) - обязательный, содержит персональные данные студентов
        /// - **contract** (Договор) - обязательный, содержит информацию о договорах на обучение
        /// - **journal** (Журнал) - обязательный, содержит номера зачетных книжек и группы
        /// - **statement** (Ведомость) - обязательный, содержит оценки по дисциплинам
        /// - **plan** (Учебный план) - обязательный, содержит учебный план с часами и типами контроля
        /// 
        /// **Процесс обработки:**
        /// 1. Из файла личных дел извлекаются базовые данные о студентах (ФИО, дата рождения, адреса и т.д.)
        /// 2. Из файла договоров добавляется информация о дате заключения договора, номере договора и приказе о зачислении
        /// 3. Из журнала добавляются номера зачетных книжек и названия групп
        /// 4. Из ведомости извлекаются оценки по дисциплинам, семестр и год обучения
        /// 5. Из учебного плана добавляются аудиторные часы, зачетные единицы и типы контроля по дисциплинам
        /// 
        /// **Требования к файлам:**
        /// - Все файлы должны быть в формате `.xlsx`
        /// - Файлы должны соответствовать ожидаемой структуре (см. документацию по форматам)
        /// - Ведомость должна содержать корректные заголовки дисциплин в строках 6-7
        /// - Учебный план должен содержать лист "ПланСвод" и листы "Курс 1", "Курс 2", "Курс 3", "Курс 4"
        /// 
        /// **Возвращаемые данные:**
        /// Возвращается полный список студентов с заполненными полями, включая:
        /// - Персональные данные
        /// - Информацию об образовании
        /// - Результаты по дисциплинам (название, оценка, семестр, год, тип контроля, часы, зачетные единицы)
        /// </remarks>
        /// <response code="200">Успешный импорт данных</response>
        /// <response code="400">Ошибка валидации файлов или неподдерживаемый формат</response>
        /// <response code="401">Требуется авторизация</response>
        /// <response code="404">Один или несколько файлов не найдены</response>
        /// <response code="500">Внутренняя ошибка сервера</response>
        [HttpPost]
        [Route("students")]
        [ProducesResponseType(typeof(IEnumerable<Student>), StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        [ProducesResponseType(StatusCodes.Status401Unauthorized)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        [ProducesResponseType(StatusCodes.Status500InternalServerError)]
        public async Task<IActionResult> ImportStudents([FromForm] ImportFilesAdapter files) 
        {
            try
            {
                var ld = files.ld;
                var contract = files.contract;
                var journal = files.journal;
                var statement = files.statement;
                var plan = files.plan;
                
                if (ld == null || ld.Length == 0 ||
                    contract == null || contract.Length == 0 ||
                    journal == null || journal.Length == 0 || 
                    statement == null || statement.Length == 0 ||
                    plan == null || plan.Length == 0) throw new ArgumentNullException("Файл не найден");
                if (System.IO.Path.GetExtension(ld.FileName) != ".xlsx" ||
                    System.IO.Path.GetExtension(contract.FileName) != ".xlsx" || 
                    System.IO.Path.GetExtension(journal.FileName) != ".xlsx" ||
                    System.IO.Path.GetExtension(statement.FileName) != ".xlsx" ||
                    System.IO.Path.GetExtension(plan.FileName) != ".xlsx") throw new FileLoadException(
                    "Неподдерживаемый формат файла. Пожалуйста, загрузите файл в формате .xlsx");
                logger.LogInformation($"Начало работы с файлом: {ld.FileName}");
                var studentsLD = await _importService.GetStudentsLD(ld);
                logger.LogInformation($"Начало работы с файлом: {contract.FileName}");
                var studentsOrder = await _importService.GetStudentsContract(contract, (List<Domain.Entities.Student>)studentsLD);
                logger.LogInformation($"Начало работы с файлом: {journal.FileName}");
                var studentsJournal = await _importService.GetStudentsJournal(journal, (List<Domain.Entities.Student>)studentsOrder);
                logger.LogInformation($"Начало работы с файлом: {statement.FileName}");
                var students = await _importService.GetStudentsStatement(statement,  (List<Domain.Entities.Student>) studentsJournal);

                logger.LogInformation($"Начало работы с файлом: {plan.FileName}");
                students = await _importService.GetStudentsEducationPlan(plan, (List<Domain.Entities.Student>)students);

                return Ok(students);
            }
            catch (ArgumentNullException ex)
            {
                logger.LogError(ex.Message);
                return NotFound(ex.Message);
            }
            catch (FileLoadException ex) 
            {
                logger.LogError(ex.Message);
                return BadRequest(ex.Message);
            }
            catch (InvalidOperationException ex)
            {
                logger.LogError(ex.Message);
                return BadRequest(ex.Message);
            }
            catch (Exception ex) 
            {
                logger.LogError(ex.Message);
                return StatusCode(500, "Ошибка на стороне сервера, обратитесь к администратору");
            }
        }
    }
}
