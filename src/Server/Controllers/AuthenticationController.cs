using Microsoft.AspNetCore.Mvc;
using static Server.Shared.Constants;
using Server.Models;
using Server.Services;
namespace Server.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class AuthenticationController : Controller
    {
        private readonly IAuthenticationService _authenticationService;

        public AuthenticationController(IAuthenticationService authenticationService) {

            _authenticationService = authenticationService;
        }

        [HttpPost("login")]
        public IActionResult Login([FromBody] LoginRequest request) {

            try
            {
                var isAuthenticated = _authenticationService.AuthenticateUser(request.UserId, request.Password);
                if (isAuthenticated)
                {
                    return Ok(new { Message = Messages.LoginSuccessful});
                }
                else
                {
                    return Unauthorized(new { Message = Messages.InvalidCredentials });
                }
            }
            catch (Exception ex)
            {
                return BadRequest(new { Message = ex.Message });
            }
        }

        [HttpPost("change-password")]
        public IActionResult ChangePassword([FromBody] ChangePasswordRequest request)
        {
            try
            {
                _authenticationService.ChangePassword(request.UserId, request.OldPassword, request.NewPassword, request.ConfirmPassword);
                return Ok(new { Message = Messages.PasswordChange });
            }
            catch (Exception ex)
            {
                return BadRequest(new { Message = ex.Message });
            }
        }
    }

 
}