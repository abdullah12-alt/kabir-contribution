using Server.Models;

namespace Server.Services;

public interface IDdsUserService
{
    Task<List<DdsUserDto>> GetAllUsersAsync();
    Task<DdsUserDto?> GetUserByIdAsync(string userId);
    Task<DdsUserDto> CreateUserAsync(CreateDdsUserRequest request, string createdBy);
    Task<DdsUserDto> UpdateUserAsync(string userId, UpdateDdsUserRequest request, string modifiedBy);
    Task<bool> DeleteUserAsync(string userId);
    Task<bool> UserExistsAsync(string userId);
    Task<bool> ValidatePasswordAsync(string userId, string password);
} 