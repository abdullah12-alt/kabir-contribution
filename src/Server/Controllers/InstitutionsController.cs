using Microsoft.AspNetCore.Mvc;
using Server.Models;
using Server.Services;
namespace Server.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class InstitutionsController : ControllerBase
    {
        private readonly IInstitutionService _service;
        private readonly ILogger<InstitutionsController> _logger;
        public InstitutionsController(IInstitutionService service, ILogger<InstitutionsController> logger)
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
                _logger.LogInformation("Returned {Count} institutions", result.Count());
                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to retrieve institutions");
                return StatusCode(500, "An error occurred while fetching data.");
            }
        }

        [HttpGet("{id}")]
        public async Task<IActionResult> Get(long id)
        {
            try
            {
                var result = await _service.GetByIdAsync(id);
                if (result == null)
                {
                    _logger.LogWarning("Institution ID={Id} not found", id);
                    return NotFound();
                }
                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to retrieve institution ID={Id}", id);
                return StatusCode(500, "An error occurred while fetching the record.");
            }
        }


        [HttpPost]
        public async Task<IActionResult> CreateOrUpdate([FromBody] Institution model, [FromQuery] string updateMode, [FromQuery] string userId)
        {
            try
            {
                var result = await _service.InsertOrUpdateAsync(model, updateMode, userId);
                if (result != 0)
                {
                    _logger.LogWarning("Insert/Update failed for InstitutionId={Id} by {User}", model.INSTITUTION_ID, userId);
                    return BadRequest("Insert or update failed.");
                }

                _logger.LogInformation("Insert/Update succeeded for InstitutionId={Id} by {User}", model.INSTITUTION_ID, userId);
                return Ok();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error inserting/updating Institution ID={Id}", model.INSTITUTION_ID);
                return StatusCode(500, "Insert or update failed.");
            }
        }


        //[HttpDelete("{id}")]
        //public async Task<IActionResult> Delete(long id, [FromQuery] string userId)
        //{
        //    try
        //    {
        //        var result = await _service.DeleteAsync(id, userId);
        //        if (result != 0)
        //        {
        //            _logger.LogWarning("Delete failed for InstitutionId={Id} by {User}", id, userId);
        //            return BadRequest("Delete failed.");
        //        }

        //        _logger.LogInformation("Deleted InstitutionId={Id} by {User}", id, userId);
        //        return Ok();
        //    }
        //    catch (Exception ex)
        //    {
        //        _logger.LogError(ex, "Error deleting Institution ID={Id}", id);
        //        return StatusCode(500, "Delete failed.");
        //    }
        //}

    }
}
