using Microsoft.AspNetCore.Mvc;
using Server.Models;
using Server.Services;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Server.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class BalanceController : ControllerBase
    {
        private readonly IBalanceService _service;
        public BalanceController(IBalanceService service)
        {
            _service = service;
        }

        [HttpGet("summary")]
        public async Task<ActionResult<BalanceSummaryDto>> GetBalanceSummary()
        {
            var summary = await _service.GetBalanceSummaryAsync();
            return Ok(summary);
        }

        [HttpGet("invalid-records")]
        public async Task<ActionResult<IList<InvalidRecord>>> GetInvalidRecords()
        {
            var records = await _service.GetInvalidRecordsAsync();
            return Ok(records);
        }

        [HttpPost("insert-balance")]
        public async Task<ActionResult<int>> InsertBalance([FromBody] BalanceInsertDto dto)
        {
            var result = await _service.InsertBalanceAsync(dto);
            return Ok(result);
        }

        [HttpGet("summary-records")]
        public async Task<ActionResult<IList<SummaryRecordDto>>> GetSummaryRecords()
        {
            var records = await _service.GetSummaryRecordsAsync();
            return Ok(records);
        }
    }
}
