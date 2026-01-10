using Dekauto.Import.Service.Domain.Interfaces;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace Dekauto.Import.Service.API.Controllers
{
    /// <summary>
    /// Контроллер для получения метрик и проверки работоспособности сервиса
    /// </summary>
    /// <remarks>
    /// Этот контроллер предоставляет endpoints для мониторинга состояния сервиса и получения статистики запросов.
    /// </remarks>
    [ApiController]
    public class MetricsController : ControllerBase
    {
        private readonly IRequestMetricsService requestMetricsService;

        public MetricsController(IRequestMetricsService requestMetricsService)
        {
            this.requestMetricsService = requestMetricsService;
        }

        /// <summary>
        /// Проверка работоспособности сервиса (Health Check)
        /// </summary>
        /// <returns>Статус доступности сервиса</returns>
        /// <remarks>
        /// Простой endpoint для проверки, что сервис запущен и отвечает на запросы.
        /// Используется системами мониторинга и балансировщиками нагрузки.
        /// 
        /// **Использование:**
        /// - Мониторинг доступности сервиса
        /// - Проверка работоспособности перед выполнением основных операций
        /// - Интеграция с системами оркестрации (Kubernetes, Docker Swarm и т.д.)
        /// </remarks>
        /// <response code="200">Сервис доступен и работает</response>
        /// <response code="400">Ошибка при проверке работоспособности</response>
        [Route("healthcheck")]
        [HttpGet]
        [ProducesResponseType(typeof(bool), StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        public IActionResult HealthCheckAsync()
        {
            try
            {
                return Ok(true);
            }
            catch
            {
                return BadRequest(false);
            }
        }

        /// <summary>
        /// Получение статистики запросов за последний период
        /// </summary>
        /// <returns>Метрики запросов к сервису</returns>
        /// <remarks>
        /// Возвращает статистику по количеству запросов к сервису за определенный период времени.
        /// 
        /// **Требования:**
        /// - Требуется авторизация через Basic Authentication
        /// 
        /// **Возвращаемые данные:**
        /// Метрики включают информацию о:
        /// - Общем количестве запросов
        /// - Количестве запросов по типам (импорт, метрики и т.д.)
        /// - Временных интервалах обработки запросов
        /// 
        /// **Использование:**
        /// - Мониторинг нагрузки на сервис
        /// - Анализ использования API
        /// - Планирование масштабирования
        /// </remarks>
        /// <response code="200">Метрики успешно получены</response>
        /// <response code="401">Требуется авторизация</response>
        /// <response code="500">Ошибка при получении метрик</response>
        [Route("requests")]
        [Authorize]
        [HttpGet]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status401Unauthorized)]
        [ProducesResponseType(StatusCodes.Status500InternalServerError)]
        public IActionResult RequestsPerPeriod()
        {
            try
            {
                return Ok(requestMetricsService.GetRecentCounters());
            }
            catch
            {
                return StatusCode(StatusCodes.Status500InternalServerError, "Не удалось получить количество запросов.");
            }
        }

    }
}

