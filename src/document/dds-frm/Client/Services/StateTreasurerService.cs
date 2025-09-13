using System.Net.Http.Json;
using Client.Models;

namespace Client.Services;

public interface IStateTreasurerService
{
    Task<IReadOnlyList<DsnItem>> GetPriorDsnsAsync();
    Task<StateTreasurerStatusDto> GetStatusAsync(DateTime postedDate);
    Task<DailyTotals> GetTotalsAsync(DateTime postedDate);
    Task<IReadOnlyList<InstitutionTotal>> GetInstitutionTotalsAsync(DateTime postedDate);
    Task<AlogResponse> GetAlogAsync(DateTime postedDate, string? sequenceNum = null, DateTime? processDate = null);
    Task<FileGenerationResponse> GenerateFilesAsync(FileGenerationRequest request);
    Task<ProcessResponse> ProcessAsync(ProcessRequest request);
}

public class StateTreasurerService : IStateTreasurerService
{
    private readonly HttpClient _httpClient;
    private readonly ILogger<StateTreasurerService> _logger;

    public StateTreasurerService(HttpClient httpClient, ILogger<StateTreasurerService> logger)
    {
        _httpClient = httpClient;
        _logger = logger;
    }

    public async Task<IReadOnlyList<DsnItem>> GetPriorDsnsAsync()
    {
        try
        {
            var response = await _httpClient.GetAsync("api/StateTreasurer/dsns");
            response.EnsureSuccessStatusCode();
            var result = await response.Content.ReadFromJsonAsync<List<DsnItem>>();
            return result ?? new List<DsnItem>();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching prior DSNs");
            throw;
        }
    }

    public async Task<StateTreasurerStatusDto> GetStatusAsync(DateTime postedDate)
    {
        try
        {
            var dateParam = postedDate.ToString("yyyy-MM-dd");
            var response = await _httpClient.GetAsync($"api/StateTreasurer/status?date={dateParam}");
            response.EnsureSuccessStatusCode();
            var result = await response.Content.ReadFromJsonAsync<StateTreasurerStatusDto>();
            return result ?? new StateTreasurerStatusDto();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching status for date {Date}", postedDate);
            throw;
        }
    }

    public async Task<DailyTotals> GetTotalsAsync(DateTime postedDate)
    {
        try
        {
            var dateParam = postedDate.ToString("yyyy-MM-dd");
            var response = await _httpClient.GetAsync($"api/StateTreasurer/totals?date={dateParam}");
            response.EnsureSuccessStatusCode();
            var result = await response.Content.ReadFromJsonAsync<DailyTotals>();
            return result ?? new DailyTotals();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching totals for date {Date}", postedDate);
            throw;
        }
    }

    public async Task<IReadOnlyList<InstitutionTotal>> GetInstitutionTotalsAsync(DateTime postedDate)
    {
        try
        {
            var dateParam = postedDate.ToString("yyyy-MM-dd");
            var response = await _httpClient.GetAsync($"api/StateTreasurer/institutions?date={dateParam}");
            response.EnsureSuccessStatusCode();
            var result = await response.Content.ReadFromJsonAsync<List<InstitutionTotal>>();
            return result ?? new List<InstitutionTotal>();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching institution totals for date {Date}", postedDate);
            throw;
        }
    }

    public async Task<AlogResponse> GetAlogAsync(DateTime postedDate, string? sequenceNum = null, DateTime? processDate = null)
    {
        try
        {
            var dateParam = postedDate.ToString("yyyy-MM-dd");
            var url = $"api/StateTreasurer/alog?date={dateParam}";
            
            if (!string.IsNullOrEmpty(sequenceNum))
                url += $"&sequenceNum={sequenceNum}";
            
            if (processDate.HasValue)
                url += $"&processDate={processDate.Value:yyyy-MM-dd}";

            var response = await _httpClient.GetAsync(url);
            response.EnsureSuccessStatusCode();
            var result = await response.Content.ReadFromJsonAsync<AlogResponse>();
            return result ?? new AlogResponse();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching ALOG for date {Date}", postedDate);
            throw;
        }
    }

    public async Task<FileGenerationResponse> GenerateFilesAsync(FileGenerationRequest request)
    {
        try
        {
            var response = await _httpClient.PostAsJsonAsync("api/StateTreasurer/generate-files", request);
            response.EnsureSuccessStatusCode();
            var result = await response.Content.ReadFromJsonAsync<FileGenerationResponse>();
            return result ?? new FileGenerationResponse();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error generating files for date {Date}", request.PostedDate);
            throw;
        }
    }

    public async Task<ProcessResponse> ProcessAsync(ProcessRequest request)
    {
        try
        {
            var response = await _httpClient.PostAsJsonAsync("api/StateTreasurer/process", request);
            response.EnsureSuccessStatusCode();
            var result = await response.Content.ReadFromJsonAsync<ProcessResponse>();
            return result ?? new ProcessResponse();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error processing request for date {Date}", request.Dsn.PostedDate);
            throw;
        }
    }
}
