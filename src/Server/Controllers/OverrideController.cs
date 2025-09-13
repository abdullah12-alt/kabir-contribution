using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Server.Models;
using Server.Services;

namespace Server.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class OverrideController : ControllerBase
    {
        private readonly IOverrideService _overrideService;
        private readonly ILogger<OverrideController> _logger;

        public OverrideController(IOverrideService overrideService, ILogger<OverrideController> logger)
        {
            _overrideService = overrideService;
            _logger = logger;
        }

        /// <summary>
        /// Validates account information for an invalid record
        /// </summary>
        /// <param name="id">Invalid record ID</param>
        /// <param name="request">Override request with account details</param>
        /// <returns>Validation result with record details</returns>
        [HttpPost("validate/{id}")]
        public async Task<IActionResult> ValidateAccount(long id, [FromBody] OverrideRequest request)
        {
            try
            {
                if (!ModelState.IsValid)
                {
                    return BadRequest(ModelState);
                }

                var result = await _overrideService.ValidateAccountAsync(id, request);

                if (result.Success)
                {
                    _logger.LogInformation("Account validation successful for InvalidRecordId={Id}", id);
                    return Ok(result);
                }
                else
                {
                    _logger.LogWarning("Account validation failed for InvalidRecordId={Id}: {Message}", id, result.Message);
                    return BadRequest(result);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to validate account for InvalidRecordId={Id}", id);
                return StatusCode(500, "Account validation failed.");
            }
        }

        /// <summary>
        /// Moves an invalid record to valid records table
        /// </summary>
        /// <param name="id">Invalid record ID</param>
        /// <param name="request">Override request with account details</param>
        /// <returns>Success or failure response</returns>
        [HttpPost("move/{id}")]
        public async Task<IActionResult> MoveToValid(long id, [FromBody] OverrideRequest request)
        {
            try
            {
                if (!ModelState.IsValid)
                {
                    return BadRequest(ModelState);
                }

                var userId = "System";

                await _overrideService.MoveRecordToValidAsync(id, request, userId);

                _logger.LogInformation("Successfully moved InvalidRecordId={Id} to valid by {User}", id, userId);
                return Ok(new { message = "Record moved to valid successfully" });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to move InvalidRecordId={Id} to valid", id);
                return StatusCode(500, "Move to valid failed.");
            }
        }

        /// <summary>
        /// Gets an invalid record by ID
        /// </summary>
        /// <param name="id">Invalid record ID</param>
        /// <returns>Invalid record details</returns>
        [HttpGet("{id}")]
        public async Task<IActionResult> GetInvalidRecord(long id)
        {
            try
            {
                var record = await _overrideService.GetInvalidRecordAsync(id);

                if (record == null)
                {
                    return NotFound($"Invalid record with ID {id} not found");
                }

                return Ok(record);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to get InvalidRecordId={Id}", id);
                return StatusCode(500, "Failed to retrieve invalid record.");
            }
        }

        /// <summary>
        /// Health check endpoint
        /// </summary>
        /// <returns>Health status</returns>
        [HttpGet("health")]
        public ActionResult<string> Health()
        {
            return Ok("Override API is running");
        }
    }
}
