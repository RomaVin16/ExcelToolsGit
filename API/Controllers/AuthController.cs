using API.Models;
using APILib.Contracts;
using APILib.Repository;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;

namespace API.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class AuthController : Controller
    {
        private readonly ISmsService _smsService;
        private readonly IAuthService _authService;
        private readonly IRepositoryContext _context;
        private readonly UserManager<ApplicationUser> _userManager;
        public AuthController(ISmsService smsService, IRepositoryContext context, UserManager<ApplicationUser> userManager, IAuthService authService)
        {
            _smsService = smsService;
            _context = context;
            _userManager = userManager;
            _authService = authService;
        }

        /// <summary>
        /// Endpoint для генерации кода
        /// </summary>
        /// <param name="phoneNumber"></param>
        /// <returns></returns>
        /// <summary>
        [HttpPost("send-otp")]
        public async Task<IActionResult> SendOtp([FromBody] string phoneNumber)
        {
            var code = _authService.GenerateSecureCode();

            try
            {
                await _authService.SaveOtpCode(phoneNumber, code);
                await _smsService.SendAsync(phoneNumber, $"Ваш код: {code}");
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Ошибка сервера: {ex}");
            }

            return Ok(new { expiresIn = 300 }); 
        }

        /// <summary>
        /// Endpoint для входа по коду 
        /// </summary>
        /// <param name="dto"></param>
        /// <returns></returns>
        [HttpPost("verify-otp")]
        public async Task<IActionResult> VerifyOtp([FromBody] VerifyOtpDto dto)
        {
            var otp = await _context.GetValidOtpCodeAsync(dto.PhoneNumber, dto.Code);
            if (otp == null || otp.ExpiryTime < DateTime.UtcNow) return Unauthorized("Неверный или просроченный код");

            var user = await _userManager.FindByNameAsync(dto.PhoneNumber);

            if (user == null)
            {
                user = new ApplicationUser
                {
                    PhoneNumber = dto.PhoneNumber,
                    UserName = dto.PhoneNumber,
                    NormalizedUserName = dto.PhoneNumber.ToUpper(), 
                    Email = $"{dto.PhoneNumber}@temp.com",
                    NormalizedEmail = $"{dto.PhoneNumber}@temp.com".ToUpper(),
                    EmailConfirmed = false, 
                    PhoneNumberConfirmed = true,
                    TwoFactorEnabled = false,
                    LockoutEnabled = true,
                    AccessFailedCount = 0,
                    SecurityStamp = Guid.NewGuid().ToString(),
                    ConcurrencyStamp = Guid.NewGuid().ToString(),
                    PasswordHash = ""
                };

                var result = await _userManager.CreateAsync(user);
                if (!result.Succeeded)
                {
                    return BadRequest(result.Errors);
                }
            }

            var token = _authService.GenerateJwtToken(user);

            return Ok(new
            {
	            Token = token,
	            User = new
	            {
		            user.Id,
		            user.UserName,
		            user.PhoneNumber
	            }
            });
		}
    }
}
