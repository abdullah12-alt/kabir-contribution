using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Server.Services;

namespace Server.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class LookupController : ControllerBase
    {
        private readonly ILookupService _lookupService;

        public LookupController(ILookupService lookupService)
        {
            _lookupService = lookupService;
        }

        [HttpGet("income-source-types")]
        public async Task<IActionResult> GetIncomeSourceTypes()
        {
            try
            {
                var result = await _lookupService.GetIncomeSourceTypesAsync();
                return Ok(result);
            }
            catch (Exception ex)
            {
                // Optional: log the error or handle more gracefully
                return StatusCode(500, "An error occurred while fetching data.");
            }
        }

    }
}
