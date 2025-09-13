using Microsoft.AspNetCore.Mvc;
using Server.Infrastructure.Logging;
using Server.Models;
using Server.Services;

namespace Server.Controllers;

[ApiController]
[Route("api/[controller]")]
public class StateTreasurerController : ControllerBase
{
    private readonly IStateTreasurerService _service;
    private readonly ILogger<StateTreasurerController> _logger;

    public StateTreasurerController(IStateTreasurerService service, ILogger<StateTreasurerController> logger)
    {
        _service = service;
        _logger = logger;
    }

    [HttpGet("status")] // api/StateTreasurer/status?date=2025-01-31
    public async Task<ActionResult<StateTreasurerStatusDto>> GetStatus([FromQuery] DateTime date)
    {
        try
        {
            if (date == default) return BadRequest("date is required");
            var status = await _service.GetStatusAsync(date);
            return Ok(status);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error getting status for {Date}", date);
            return StatusCode(500, "Error fetching status");
        }
    }

    [HttpGet("dsns")] // api/StateTreasurer/dsns
    public async Task<ActionResult<IReadOnlyList<DsnItem>>> GetPriorDsns()
    {
        try
        {
            var dsns = await _service.GetPriorDsnsAsync();
            return Ok(dsns);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching prior DSNs");
            return StatusCode(500, "Error fetching DSNs");
        }
    }



    [HttpGet("totals")] // api/StateTreasurer/totals?date=2025-01-31
    public async Task<ActionResult<DailyTotals>> GetTotals([FromQuery] DateTime date)
    {
        try
        {
            if (date == default) return BadRequest("date is required");
            var totals = await _service.GetTotalsAsync(date);
            return Ok(totals);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching totals for {Date}", date);
            return StatusCode(500, "Error fetching totals");
        }
    }

    [HttpGet("institutions")] // api/StateTreasurer/institutions?date=2025-01-31
    public async Task<ActionResult<IReadOnlyList<InstitutionTotal>>> GetInstitutionTotals([FromQuery] DateTime date)
    {
        try
        {
            if (date == default) return BadRequest("date is required");
            var rows = await _service.GetInstitutionTotalsAsync(date);
            return Ok(rows);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching institution totals for {Date}", date);
            return StatusCode(500, "Error fetching institution totals");
        }
    }

    [HttpGet("alog")] // api/StateTreasurer/alog?date=2025-01-31
    public async Task<ActionResult<AlogResponse>> GetAlog([FromQuery] DateTime date, [FromQuery] string? sequenceNum, [FromQuery] DateTime? processDate)
    {
        try
        {
            if (date == default) return BadRequest("date is required");
            var resp = await _service.BuildAlogAsync(date, sequenceNum, processDate);
            return Ok(resp);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error building alog for {Date}", date);
            return StatusCode(500, "Error building alog");
        }
    }

    [HttpPost("process")] // api/StateTreasurer/process
    public async Task<ActionResult<ProcessResponse>> Process([FromBody] ProcessRequest request)
    {
        if (!ModelState.IsValid)
            return BadRequest(ModelState);

        try
        {
            var createdBy = "SYSTEM"; // TODO: replace with authenticated user
            var userId = "SYSTEM";
            // 1. Insert DSN
            var dsnResult = await _service.InsertDsnAsync(request.Dsn, request.Dsn.PostedDate, createdBy);
            if (dsnResult != 0)
                return StatusCode(500, $"Insert DSN failed with code {dsnResult}");

            // 2. Generate files
            var fileReq = new FileGenerationRequest
            {
                PostedDate = request.Dsn.PostedDate,
                ProcessDate = DateTime.Now,
                DepSeqNum = request.Dsn.DepSeqNum
            };
            var files = await _service.GenerateFilesAsync(fileReq);

            // 3. Send email
            var emailReq = new EmailRequest
            {
                PostedDate = request.Dsn.PostedDate,
                DepSeqNum = request.Dsn.DepSeqNum,
                SendAlogs = request.SendAlogs
            };
            var emailResp = await _service.SendEmailAsync(emailReq, userId);

            return Ok(new ProcessResponse
            {
                Success = emailResp.Success,
                Message = emailResp.Success
                    ? "Process completed successfully"
                    : $"Email step failed: {emailResp.Message}",
                Files = files,
                EmailResponse = emailResp
            });
        }
        catch (InvalidOperationException dupEx)
        {
            return Conflict(new { message = dupEx.Message });
        }
        catch (ArgumentException argEx)
        {
            return BadRequest(new { message = argEx.Message });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error running State Treasurer process");
            return StatusCode(500, "Error running process");
        }
    }

    [HttpPost("generate-files")] // api/StateTreasurer/generate-files
    public async Task<ActionResult<FileGenerationResponse>> GenerateFiles([FromBody] FileGenerationRequest request)
    {
        if (!ModelState.IsValid)
            return BadRequest(ModelState);

        try
        {
            var files = await _service.GenerateFilesAsync(request);
            return Ok(files);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error generating files for {Date}", request.PostedDate);
            return StatusCode(500, "Error generating files");
        }
    }

    [HttpGet("download-pa-file")] // api/StateTreasurer/download-pa-file?date=2025-01-31&processDate=2025-01-31
    public async Task<IActionResult> DownloadPaFile([FromQuery] DateTime date, [FromQuery] DateTime processDate)
    {
        try
        {
            if (date == default || processDate == default) return BadRequest("date and processDate are required");

            var request = new FileGenerationRequest
            {
                PostedDate = date,
                ProcessDate = processDate
            };

            var response = await _service.GenerateFilesAsync(request);
            var bytes = System.Text.Encoding.UTF8.GetBytes(response.PaFileContent);

            return File(bytes, "text/plain", response.PaFileName);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error downloading PA file for {Date}", date);
            return StatusCode(500, "Error generating PA file");
        }
    }

    [HttpGet("download-pf-file")] // api/StateTreasurer/download-pf-file?date=2025-01-31&processDate=2025-01-31
    public async Task<IActionResult> DownloadPfFile([FromQuery] DateTime date, [FromQuery] DateTime processDate)
    {
        try
        {
            if (date == default || processDate == default) return BadRequest("date and processDate are required");

            var request = new FileGenerationRequest
            {
                PostedDate = date,
                ProcessDate = processDate
            };

            var response = await _service.GenerateFilesAsync(request);
            var bytes = System.Text.Encoding.UTF8.GetBytes(response.PfFileContent);

            return File(bytes, "text/plain", response.PfFileName);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error downloading PF file for {Date}", date);
            return StatusCode(500, "Error generating PF file");
        }
    }


}

