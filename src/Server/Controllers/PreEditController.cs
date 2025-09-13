using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Server.Models;
using Server.Services;

namespace Server.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class PreEditController : ControllerBase
    {
        private readonly IPreEditService _service;
        private readonly ILogger<PreEditController> _logger;

        public PreEditController(IPreEditService service, ILogger<PreEditController> logger)
        {
            _service = service;
            _logger = logger;
        }

        [HttpGet("{baiFileId}")]
        public async Task<IActionResult> GetInvalidRecords([FromRoute]int baiFileId)
        {
            try
            {
                var records = await _service.GetInvalidRecordsByIdAsync(baiFileId);
                _logger.LogInformation("Returned {Count} invalid records", records.Count());
                return Ok(records);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to retrieve invalid records");
                return StatusCode(500, "An error occurred while fetching records.");
            }
        }

        [HttpGet()]
        public async Task<IActionResult> GetAllInvalidRecords()
        {
            try
            {
                var records = await _service.GetAllInvalidRecordsAsync();
                _logger.LogInformation("Returned {Count} invalid records", records.Count());
                return Ok(records);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to retrieve invalid records");
                return StatusCode(500, "An error occurred while fetching records.");
            }
        }


        [HttpPut]
        public async Task<IActionResult> Update([FromBody] InvalidRecord record)
        {
            try
            {
                await _service.UpdateInvalidRecordAsync(record);
                _logger.LogInformation("Updated InvalidRecordId={Id} by User={User}", record.INVALID_RECORD_ID, record.LAST_MOD_BY);
                return NoContent();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to update InvalidRecordId={Id}", record.INVALID_RECORD_ID);
                return StatusCode(500, "Update failed.");
            }
        }

        [HttpPost("move/{id}")]
        public async Task<IActionResult> MoveToValid(long id)
        {
            try
            {
                await _service.MoveRecordToValidAsync(id);
                _logger.LogInformation("Moved InvalidRecordId={Id}", id);
                return Ok();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to move InvalidRecordId={Id} to valid", id);
                return StatusCode(500, "Move to valid failed.");
            }
        }

        [HttpPost("recoup")]
        public async Task<IActionResult> Recoup([FromQuery] long creditId, long debitId, string userId)
        {
            try
            {
                await _service.RecoupAsync(creditId, debitId, userId);
                _logger.LogInformation("Recouped CreditId={CreditId}, DebitId={DebitId} by {User}", creditId, debitId, userId);
                return Ok();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to recoup CreditId={CreditId} and DebitId={DebitId}", creditId, debitId);
                return StatusCode(500, "Recoup failed.");
            }
        }

        [HttpDelete("{id}")]
        public async Task<IActionResult> Delete(long id)
        {
            try
            {
                await _service.DeleteRecordAsync(id);
                _logger.LogInformation("Deleted InvalidRecordId={Id} by {User}", id);
                return Ok();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to delete InvalidRecordId={Id}", id);
                return StatusCode(500, "Delete failed.");
            }
        }
        [HttpGet("hidden-transactions")]
        public async Task<IActionResult> GetHiddeTransactionsAsync()
        {
            try
            {
                var records = await _service.GetHiddeTransactionsAsync();
                _logger.LogInformation("Returned {Count} invalid records", records.Count());
                return Ok(records);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to retrieve invalid records");
                return StatusCode(500, $"An error occurred while fetching records.{ex}");
            }
        }
        [HttpPost("undelete/{id}")]
        public async Task<IActionResult> Undelete(long id, [FromQuery] string userId)
        {
            try
            {
                await _service.UndeleteRecordAsync(id, userId);
                _logger.LogInformation("Undeleted InvalidRecordId={Id} by {User}", id, userId);
                return Ok();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to undelete InvalidRecordId={Id}", id);
                return StatusCode(500, "Undelete failed.");
            }
        }

        [HttpPost("hide-preedit")]
        public async Task<IActionResult> HidePreEditRecord([FromBody] HidePreEditRequestDto request)
        {
            try
            {
                var result = await _service.HidePreEditRecordAsync(request.InvalidRecordId, request.RecordStatus, request.UserId);

                if (result == 0)
                {
                    return Ok(new { success = true, message = "Record updated successfully" });
                }
                else
                {
                    return StatusCode(500, new { success = false, message = "Stored procedure returned an error", code = result });
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Exception while hiding PreEdit record");
                return StatusCode(500, new { success = false, message = "An error occurred while updating the record." });
            }
        }

    }


}
