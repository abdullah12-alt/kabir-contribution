using Microsoft.AspNetCore.Mvc;
using Server.Infrastructure.Logging;
using Server.Models;
using Server.Services;

namespace Server.Controllers;

[ApiController]
[Route("api/[controller]")]
public class DdsUsersController : ControllerBase
{
    private readonly IDdsUserService _ddsUserService;
    private readonly IAppLogger<DdsUsersController> _logger;

    public DdsUsersController(IDdsUserService ddsUserService, IAppLogger<DdsUsersController> logger)
    {
        _ddsUserService = ddsUserService;
        _logger = logger;
    }

    /// <summary>
    /// Get all DDS users
    /// </summary>
    /// <returns>List of all DDS users</returns>
    [HttpGet]
    public async Task<ActionResult<List<DdsUserDto>>> GetAllUsers()
    {
        try
        {
            var users = await _ddsUserService.GetAllUsersAsync();
            return Ok(users);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving all DDS users");
            return StatusCode(500, "An error occurred while retrieving users.");
        }
    }

    /// <summary>
    /// Get a specific DDS user by ID
    /// </summary>
    /// <param name="userId">The user ID</param>
    /// <returns>The DDS user</returns>
    [HttpGet("{userId}")]
    public async Task<ActionResult<DdsUserDto>> GetUserById(string userId)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(userId))
            {
                return BadRequest("User ID is required.");
            }

            var user = await _ddsUserService.GetUserByIdAsync(userId);
            if (user == null)
            {
                return NotFound($"User with ID '{userId}' not found.");
            }

            return Ok(user);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving DDS user: {UserId}", userId);
            return StatusCode(500, "An error occurred while retrieving the user.");
        }
    }

    /// <summary>
    /// Create a new DDS user
    /// </summary>
    /// <param name="request">The user creation request</param>
    /// <returns>The created DDS user</returns>
    [HttpPost]
    public async Task<ActionResult<DdsUserDto>> CreateUser([FromBody] CreateDdsUserRequest request)
    {
        try
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            // TODO: Get the current user from authentication context
            var createdBy = "SYSTEM"; // This should come from the authenticated user

            var createdUser = await _ddsUserService.CreateUserAsync(request, createdBy);
            return CreatedAtAction(nameof(GetUserById), new { userId = createdUser.UserId }, createdUser);
        }
        catch (InvalidOperationException ex)
        {
            _logger.LogWarning(ex, "Invalid operation while creating DDS user: {UserId}", request.UserId);
            return BadRequest(ex.Message);
        }
        catch (ArgumentException ex)
        {
            _logger.LogWarning(ex, "Invalid argument while creating DDS user: {UserId}", request.UserId);
            return BadRequest(ex.Message);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error creating DDS user: {UserId}", request.UserId);
            return StatusCode(500, "An error occurred while creating the user.");
        }
    }

    /// <summary>
    /// Update an existing DDS user
    /// </summary>
    /// <param name="userId">The user ID</param>
    /// <param name="request">The user update request</param>
    /// <returns>The updated DDS user</returns>
    [HttpPut("{userId}")]
    public async Task<ActionResult<DdsUserDto>> UpdateUser(string userId, [FromBody] UpdateDdsUserRequest request)
    {
        try
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (string.IsNullOrWhiteSpace(userId))
            {
                return BadRequest("User ID is required.");
            }

            // TODO: Get the current user from authentication context
            var modifiedBy = "SYSTEM"; // This should come from the authenticated user

            var updatedUser = await _ddsUserService.UpdateUserAsync(userId, request, modifiedBy);
            return Ok(updatedUser);
        }
        catch (InvalidOperationException ex)
        {
            _logger.LogWarning(ex, "Invalid operation while updating DDS user: {UserId}", userId);
            return BadRequest(ex.Message);
        }
        catch (ArgumentException ex)
        {
            _logger.LogWarning(ex, "Invalid argument while updating DDS user: {UserId}", userId);
            return BadRequest(ex.Message);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error updating DDS user: {UserId}", userId);
            return StatusCode(500, "An error occurred while updating the user.");
        }
    }

    /// <summary>
    /// Delete a DDS user
    /// </summary>
    /// <param name="userId">The user ID</param>
    /// <returns>Success status</returns>
    [HttpDelete("{userId}")]
    public async Task<ActionResult> DeleteUser(string userId)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(userId))
            {
                return BadRequest("User ID is required.");
            }

            var success = await _ddsUserService.DeleteUserAsync(userId);
            if (!success)
            {
                return NotFound($"User with ID '{userId}' not found.");
            }

            return NoContent();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error deleting DDS user: {UserId}", userId);
            return StatusCode(500, "An error occurred while deleting the user.");
        }
    }

    /// <summary>
    /// Check if a user exists
    /// </summary>
    /// <param name="userId">The user ID</param>
    /// <returns>True if user exists, false otherwise</returns>
    [HttpGet("{userId}/exists")]
    public async Task<ActionResult<bool>> UserExists(string userId)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(userId))
            {
                return BadRequest("User ID is required.");
            }

            var exists = await _ddsUserService.UserExistsAsync(userId);
            return Ok(exists);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error checking if DDS user exists: {UserId}", userId);
            return StatusCode(500, "An error occurred while checking if the user exists.");
        }
    }

    /// <summary>
    /// Validate a user's password
    /// </summary>
    /// <param name="userId">The user ID</param>
    /// <param name="password">The password to validate</param>
    /// <returns>True if password is valid, false otherwise</returns>
    [HttpPost("{userId}/validate-password")]
    public async Task<ActionResult<bool>> ValidatePassword(string userId, [FromBody] string password)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(userId))
            {
                return BadRequest("User ID is required.");
            }

            if (string.IsNullOrWhiteSpace(password))
            {
                return BadRequest("Password is required.");
            }

            var isValid = await _ddsUserService.ValidatePasswordAsync(userId, password);
            return Ok(isValid);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error validating password for DDS user: {UserId}", userId);
            return StatusCode(500, "An error occurred while validating the password.");
        }
    }
} 