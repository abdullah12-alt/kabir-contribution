using Microsoft.AspNetCore.Mvc;
using Server.Models;
using Server.Services;

namespace Server.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class DDConfigController : ControllerBase
    {

        private readonly IDDConfigService _service;
        private readonly ILogger<DDConfigController> _logger;

        public DDConfigController(IDDConfigService service, ILogger<DDConfigController> logger)
        {
            _service = service;
            _logger = logger;
        }
        [HttpGet]
        public async Task<IActionResult> Get()
        {
            try
            {
                var config = await _service.GetConfigAsync();
                if (config == null)
                {
                    _logger.LogWarning("DDConfig record not found");
                    return NotFound();
                }

                _logger.LogInformation("Returned DDConfig record ID={Id}", config.CONFIG_ID);
                return Ok(config);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to retrieve DDConfig record");
                return StatusCode(500, "An error occurred while fetching the configuration.");
            }
        }

        [HttpPut]
        public async Task<IActionResult> Update([FromBody] DDConfigInfo model, [FromQuery] string updateMode)
        {
            try
            {
                var result = await _service.UpdateConfigAsync(model, updateMode);

                if (result != 0)
                {
                    _logger.LogWarning("Update failed for DDConfig ID={Id}, Mode={Mode}", model.CONFIG_ID, updateMode);
                    return BadRequest("Error occurred updating the configuration.");
                }

                _logger.LogInformation("Updated DDConfig ID={Id} with mode={Mode}", model.CONFIG_ID, updateMode);
                return Ok();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating DDConfig ID={Id}", model.CONFIG_ID);
                return StatusCode(500, "Update failed.");
            }
        }
    }

}
