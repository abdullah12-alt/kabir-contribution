using Server.Infrastructure.Logging;
using Server.Models;
using Server.Repositories;

namespace Server.Services
{
    public interface IRegionService
    {
        Task<List<RegionDto>> GetAllRegionsAsync();
        Task<RegionDto?> GetRegionByIdAsync(long regionId);
        Task<long> AddRegionAsync(RegionDto region);
        Task<bool> UpdateRegionAsync(RegionDto region);
        Task<bool> DeleteRegionAsync(long regionId);
    }
    public class RegionService : IRegionService
    {
        private readonly IRegionRepository _repo;
        private readonly IAppLogger<RegionService> _logger;

        public RegionService(IRegionRepository repo, IAppLogger<RegionService> logger)
        {
            _repo = repo;
            _logger = logger;
        }

        public async Task<List<RegionDto>> GetAllRegionsAsync()
        {
            _logger.LogInformation("Getting all regions...");
            var result = await _repo.GetAllRegionsAsync();
            _logger.LogInformation("Retrieved {Count} regions.", result.Count);
            return result;
        }

        public async Task<RegionDto?> GetRegionByIdAsync(long regionId)
        {
            _logger.LogInformation("Getting region by id: {RegionId}", regionId);
            return await _repo.GetRegionByIdAsync(regionId);
        }

        public async Task<long> AddRegionAsync(RegionDto region)
        {
            _logger.LogInformation("Adding new region: {Region}", region.Region);
            return await _repo.AddRegionAsync(region);
        }

        public async Task<bool> UpdateRegionAsync(RegionDto region)
        {
            _logger.LogInformation("Updating region: {RegionId}", region.RegionId);
            return await _repo.UpdateRegionAsync(region);
        }

        public async Task<bool> DeleteRegionAsync(long regionId)
        {
            _logger.LogInformation("Deleting region: {RegionId}", regionId);
            return await _repo.DeleteRegionAsync(regionId);
        }
    }

}
