using Microsoft.AspNetCore.Mvc;
using Server.Models;
using Server.Services;

namespace Server.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class IncomeSourceTypesController : ControllerBase
    {
        private readonly IIncomeSourceTypeService _service;
        private readonly ILogger<IncomeSourceTypesController> _logger;

        public IncomeSourceTypesController(IIncomeSourceTypeService service, ILogger<IncomeSourceTypesController> logger)
        {
            _service = service;
            _logger = logger;
        }

        [HttpGet]
        public async Task<IActionResult> GetAll()
        {
            try
            {
                var result = await _service.GetAllAsync();
                _logger.LogInformation("Returned {Count} income source types", result.Count());
                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to fetch income source types");
                return StatusCode(500, "An error occurred while fetching data.");
            }
        }

        [HttpGet("{id}")]
        public async Task<IActionResult> Get(int id)
        {
            try
            {
                var result = await _service.GetByIdAsync(id);
                if (result == null)
                {
                    _logger.LogWarning("Income source type ID={Id} not found", id);
                    return NotFound();
                }
                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to fetch income source type ID={Id}", id);
                return StatusCode(500, "An error occurred while fetching the record.");
            }
        }

        [HttpPost]
        public async Task<IActionResult> CreateOrUpdate(
            [FromBody] IncomeSourceType model,
            [FromQuery] string updateMode,
            [FromQuery] string userId)
        {
            try
            {
                var result = await _service.InsertOrUpdateAsync(model, updateMode, userId);
                if (result != 0)
                {
                    _logger.LogWarning("Insert/Update failed for IncomeSourceTypeId={Id} by {User}", model.INCOME_SOURCE_TYPE_ID, userId);
                    return BadRequest("Insert or update failed.");
                }

                _logger.LogInformation("Insert/Update succeeded for IncomeSourceTypeId={Id} by {User}", model.INCOME_SOURCE_TYPE_ID, userId);
                return Ok();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error inserting/updating income source type ID={Id}", model.INCOME_SOURCE_TYPE_ID);
                return StatusCode(500, "Insert or update failed.");
            }
        }

        [HttpDelete("{id}")]
        public async Task<IActionResult> Delete(long id, [FromQuery] string userId)
        {
            try
            {
                var result = await _service.DeleteAsync(id, userId);
                if (result != 0)
                {
                    _logger.LogWarning("Delete failed for IncomeSourceTypeId={Id} by {User}", id, userId);
                    return BadRequest("Delete failed.");
                }

                _logger.LogInformation("Deleted IncomeSourceTypeId={Id} by {User}", id, userId);
                return Ok();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting income source type ID={Id}", id);
                return StatusCode(500, "Delete failed.");
            }
        }
    }
}
