using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Server.Services;

namespace Server.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class TransactionsController : ControllerBase
    {

        private readonly ITransactionService _transactionService;

        public TransactionsController(ITransactionService transactionService)
        {
            _transactionService = transactionService;
        }

        [HttpGet("unvalidated")]
        public async Task<IActionResult> GetAll()
        {
            var data = await _transactionService.GetAllTransactionsAsync();
            return Ok(data);
        }

        [HttpGet("invalid")]
        public async Task<IActionResult> GetInvalid()
        {
            var data = await _transactionService.GetInvalidTransactionsAsync();
            return Ok(data);
        }

        [HttpGet("valid")]
        public async Task<IActionResult> GetValid()
        {
            var data = await _transactionService.GetValidTransactionsAsync();
            return Ok(data);
        }
    }
}
