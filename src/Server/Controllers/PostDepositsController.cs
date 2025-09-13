// PostDepositsController.cs
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using Server.Models;
using Server.Services;
using Server.Shared;
using System.Threading.Tasks;

namespace Server.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class PostDepositsController : ControllerBase
    {
        private readonly IPostDepositsService _service;
        private readonly PostingConfig _config;

        public PostDepositsController(IPostDepositsService service, IOptions<PostingConfig> config)
        {
            _service = service;
            _config = config.Value;

        }

        /// <summary>
        /// Prepares and creates HL7 files for all valid records, ready for Affinity transfer.
        /// </summary>
        /// <param name="userId">User performing the operation</param>
        /// <param name="flatFileDirectory">Directory to save flat files</param>
        /// <param name="processingId">Processing ID ("T" for test, "P" for prod)</param>
        [HttpPost("post")]
        public async Task<IActionResult> PostDeposits()
        {
            string processingId = _config.ProcessingId;
            string userId = _config.UserId;
            string flatFileDirectory = _config.FlatFileDirectory;
            var result = await _service.PostDepositsAsync(userId, flatFileDirectory, processingId);
            if (!result.Success) 
                return BadRequest(result.Message);

            return Ok(result);
        }

        [HttpGet("counts")]
        public async Task<ActionResult<PostCounts>> GetPostCounts()
        {
            try
            {
                var counts = await _service.GetPostCountsAsync();
                return Ok(counts);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }
    }
}