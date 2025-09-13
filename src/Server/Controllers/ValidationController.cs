using Microsoft.AspNetCore.Mvc;
using Server.Services;
using System.Threading.Tasks;
using static Server.Shared.Constants;
using Server.Models;
using Server.Infrastructure.Logging;
namespace Server.Controllers {

    [Route("api/validation")]
    [ApiController]

    public class ValidationController : ControllerBase {
        private readonly IValidationService _validationService;
        private readonly IAppLogger<ValidationController> _logger;


        public ValidationController(IValidationService validationService, IAppLogger<ValidationController> logger)
        {

            _validationService = validationService;
            _logger = logger;

        }
        /// <summary>
        /// Validates the bank file records based on the given BAI file ID.
        /// </summary>
        
        [HttpPost()]
        public async Task<IActionResult> ValidateRecords() {
            _logger.LogInformation("Validation requested");


            var result = await _validationService.ValidateRecordsAsync();

            if (result.success) {
                _logger.LogInformation("Validation completed successfully");

                return Ok(new { message = Messages.ValidationSuccess, result });

            } 
            else {
                _logger.LogWarning("Validation failed", new {  errors = result.errors });

                return BadRequest(new { message = Messages.ValidationFailed ,errors = result.errors });

            }
        }
    }
}
