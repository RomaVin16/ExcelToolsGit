using API.Models;
using APILib.Contracts;
using APILib.Repository;
using Microsoft.IdentityModel.Tokens;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Security.Cryptography;
using System.Text;

namespace APILib.Services
{
    public class AuthService: IAuthService
    {
        private readonly IRepositoryContext _context;
        private readonly JwtSettings _jwtSettings;

        public AuthService(IRepositoryContext repositoryContext, JwtSettings jwtSettings)
        {
            _context = repositoryContext;
            _jwtSettings = jwtSettings;
        }

        public string GenerateJwtToken(ApplicationUser user)
        {
            if (user == null)
                throw new ArgumentNullException(nameof(user));

            var tokenHandler = new JwtSecurityTokenHandler();
            var key = Encoding.UTF8.GetBytes(_jwtSettings.Secret);

            var userId = user.Id;

			var claims = new List<Claim>
            {
                new Claim(ClaimTypes.NameIdentifier, user.Id), 
                new Claim(ClaimTypes.Name, user.UserName),
                new Claim(JwtRegisteredClaimNames.Sub, userId), 
                new Claim(JwtRegisteredClaimNames.Jti, Guid.NewGuid().ToString()),
				new Claim(ClaimTypes.MobilePhone, user.PhoneNumber),
                new Claim(ClaimTypes.Email, user.Email)
            };

            var tokenDescriptor = new SecurityTokenDescriptor
            {
                Subject = new ClaimsIdentity(claims),
                Expires = DateTime.UtcNow.AddHours(1),
                SigningCredentials = new SigningCredentials(new SymmetricSecurityKey(key), SecurityAlgorithms.HmacSha256Signature)
            };

            var token = tokenHandler.CreateToken(tokenDescriptor);
            return tokenHandler.WriteToken(token);
        }


        /// <summary>
        /// Генерация криптобезопасного кода
        /// </summary>
        /// <returns></returns>
        public string GenerateSecureCode()
        {
            using var rng = RandomNumberGenerator.Create();
            var bytes = new byte[4];
            rng.GetBytes(bytes);
            var number = BitConverter.ToUInt32(bytes) % 1_000_000;
            return number.ToString("D6");
        }

        public async Task SaveOtpCode(string phoneNumber, string code)
        {
            await _context.DeleteAllForPhoneAsync(phoneNumber);

            var otp = new OtpCode
            {
                PhoneNumber = phoneNumber,
                Code = code,
                ExpiryTime = DateTime.UtcNow.AddMinutes(5),
            };

            await _context.CreateOtpCodeAsync(otp);
        }

    }
}
