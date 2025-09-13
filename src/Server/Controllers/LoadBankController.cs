using Microsoft.AspNetCore.Mvc;
using Server.Infrastructure.Logging;
using Server.Models;
using Server.Services;
using Server.Shared.BaiFileContext;

namespace DirectDepositSystem.Controllers
{
    [ApiController]
    [Route("api/bank-files")]

    public class LoadBankController : ControllerBase
    {
        private readonly ILoadBankService _loadBankService;
        private readonly IValidationService _validationService;
        private readonly IBaiFileContext _baiFileContext;
        private readonly IAppLogger<LoadBankController> _logger;

        private readonly string[] _allowedBaiExtensions = { ".bai2", ".txt" };
        private readonly string[] _allowedDetailExtensions = { ".csv" };
        public LoadBankController(
              ILoadBankService loadBankService,
              IValidationService validationService,
              IAppLogger<LoadBankController> logger,
              IBaiFileContext baiFileContext)
        {
            _loadBankService = loadBankService;
            _logger = logger;
            _validationService = validationService;
            _baiFileContext = baiFileContext;
        }

        /// <summary>
        /// Uploads a bank file for processing.
        /// </summary>
        [HttpPost("load")]
        [ProducesResponseType(typeof(ApiResponse), StatusCodes.Status200OK)]
        [ProducesResponseType(typeof(ApiResponse), StatusCodes.Status400BadRequest)]
        [ProducesResponseType(typeof(ApiResponse), StatusCodes.Status500InternalServerError)]
        public async Task<ActionResult<ApiResponse>> UploadFiles([FromForm] LoadFileRequest request)
        {
            _logger.LogInformation("UploadFiles called with {@Request}", new
            {
                BaiFile = request?.BaiFile?.FileName,
                DetailFile = request?.DetailFile?.FileName
            });

                // 1. Validate files
            if (request?.BaiFile == null || request.DetailFile == null)
            {
                return BadRequest(new ApiResponse(false, "Files must be provided."));
            }
            if (!IsValidUpload(request?.BaiFile!, request?.DetailFile!))
            {
                return BadRequest(new ApiResponse(false, "Invalid files provided"));
            }

            try
            {
                // 2. Process Files
                var processResult = await _loadBankService.ProcessFiles(
                    request?.BaiFile!,
                    request?.DetailFile!,
                    userId: "1" // TODO: replace with actual user context
                );

                if (!processResult.Success)
                {
                    _logger.LogWarning("File processing failed: {Message}", processResult.Message);
                    return BadRequest(new ApiResponse(false, processResult.Message));
                }

                _logger.LogInformation("Files passed extension validation", new
                {
                    BaiFile = request?.BaiFile?.FileName,
                    DetailFile = request?.DetailFile?.FileName
                });

                // 3. Continue only if BaiFileId exists
                if (!string.IsNullOrEmpty(processResult.BaiFileId))
                {
                    var result = await _validationService.ValidateRecordsAsync();
                    _logger.LogWarning("File processing failed", processResult);
                    _baiFileContext.BaiFileId = processResult.BaiFileId;

                    return Ok(new ApiResponse(
                        true,
                        "Files processed successfully",
                        new
                        {
                            FileName = request?.BaiFile?.Name,
                            Error = processResult
                        }
                    ));
                }

                // 4. BaiFileId is missing unexpectedly
                return StatusCode(500, new ApiResponse(false, "Processing succeeded, but file ID is missing."));
            }
            catch (Exception ex)
            {

                _logger.LogError("Error during load", ex, request);
                return StatusCode(500, new ApiResponse(false, $"Error processing file: {ex.Message}"));
            }
        }

        // <summary>
        /// Checks whether a work file exists.
        /// </summary>
        [HttpGet("workfile-exists")]
        [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status500InternalServerError)]
        public async Task<IActionResult> WorkFileExists()
        {
            try
            {
                var exists = await _loadBankService.AnyWorkFileRecordsAsync();
                return Ok(new { exists });
            }
            catch (Exception ex)
            {
                _logger.LogError("Error checking work file records", ex);
                return StatusCode(500, "An error occurred while checking data.");
            }
        }


        // <summary>
        /// helper
        // </summary>

        private bool IsValidUpload(IFormFile baiFile, IFormFile detailFile)
        {
            if (baiFile == null || baiFile.Length == 0 ||
                detailFile == null || detailFile.Length == 0)
            {
                return false;
            }

            var baiExt = Path.GetExtension(baiFile.FileName).ToLowerInvariant();
            var detailExt = Path.GetExtension(detailFile.FileName).ToLowerInvariant();

            if (!_allowedBaiExtensions.Contains(baiExt))
            {

                _logger.LogWarning($"Invalid BAI file extension: {baiExt}");
                return false;
            }

            if (!_allowedDetailExtensions.Contains(detailExt))
            {
                _logger.LogWarning($"Invalid detail file extension: {detailExt}");
                return false;
            }

            return true;
        }


    }

}

