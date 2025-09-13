namespace Server.Models
{
    public class LoginRequest
    {
        public required string UserId { get; set; }
        public required string Password { get; set; }
    }

    public class ChangePasswordRequest
    {
        public required string UserId { get; set; }
        public required string OldPassword { get; set; }
        public required string NewPassword { get; set; }
        public required string ConfirmPassword { get; set; }
    }
}
