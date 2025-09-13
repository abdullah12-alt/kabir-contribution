using System.ComponentModel.DataAnnotations;

namespace Server.Models;

public class DdsUserDto
{
    public string UserId { get; set; } = string.Empty;
    
    [Required]
    [MaxLength(50)]
    public string UserLastName { get; set; } = string.Empty;
    
    [Required]
    [MaxLength(50)]
    public string UserFirstName { get; set; } = string.Empty;
    
    public string? Password { get; set; }
    
    public string? PasswordConfirm { get; set; }
    
    public string RecordStatus { get; set; } = "A"; // A = Active, I = Inactive
    
    public string? CreatedBy { get; set; }
    
    public DateTime? CreatedDateTime { get; set; }
    
    public string? LastModBy { get; set; }
    
    public DateTime? LastModDateTime { get; set; }
    
    public bool IsDisabled => RecordStatus == "I";
    
    public string FullName => $"{UserLastName}, {UserFirstName}".Trim();
}

public class CreateDdsUserRequest
{
    [Required]
    [MaxLength(10)]
    public string UserId { get; set; } = string.Empty;
    
    [Required]
    [MaxLength(50)]
    public string UserLastName { get; set; } = string.Empty;
    
    [Required]
    [MaxLength(50)]
    public string UserFirstName { get; set; } = string.Empty;
    
    [Required]
    [MinLength(6)]
    public string Password { get; set; } = string.Empty;
    
    [Required]
    [Compare("Password", ErrorMessage = "Password and confirmation password do not match.")]
    public string PasswordConfirm { get; set; } = string.Empty;
    
    public bool IsDisabled { get; set; } = false;
}

public class UpdateDdsUserRequest
{
    [Required]
    [MaxLength(50)]
    public string UserLastName { get; set; } = string.Empty;
    
    [Required]
    [MaxLength(50)]
    public string UserFirstName { get; set; } = string.Empty;
    
    [MinLength(6)]
    public string? Password { get; set; }
    
    [Compare("Password", ErrorMessage = "Password and confirmation password do not match.")]
    public string? PasswordConfirm { get; set; }
    
    public bool IsDisabled { get; set; } = false;
} 