// Controllers/RegionController.cs
using Microsoft.AspNetCore.Mvc;
using Server.Models;
using Server.Services;

[ApiController]
[Route("api/[controller]")]
public class RegionController : ControllerBase
{
    private readonly IRegionService _service;

    public RegionController(IRegionService service)
    {
        _service = service;
    }

    [HttpGet]
    public async Task<ActionResult<List<RegionDto>>> GetAll()
    {
        var regions = await _service.GetAllRegionsAsync();
        return Ok(regions);
    }

    [HttpGet("{id}")]
    public async Task<ActionResult<RegionDto>> Get(long id)
    {
        var region = await _service.GetRegionByIdAsync(id);
        if (region == null) return NotFound();
        return Ok(region);
    }

    [HttpPost]
    public async Task<ActionResult> Create([FromBody] RegionDto region)
    {
        var id = await _service.AddRegionAsync(region);
        return CreatedAtAction(nameof(Get), new { id }, region);
    }

    [HttpPut("{id}")]
    public async Task<ActionResult> Update(long id, [FromBody] RegionDto region)
    {
        if (id != region.RegionId) return BadRequest();
        var success = await _service.UpdateRegionAsync(region);
        if (!success) return NotFound();
        return NoContent();
    }

    [HttpDelete("{id}")]
    public async Task<ActionResult> Delete(long id)
    {
        var success = await _service.DeleteRegionAsync(id);
        if (!success) return NotFound();
        return NoContent();
    }
}