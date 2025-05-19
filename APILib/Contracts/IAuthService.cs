using APILib.Repository;

namespace APILib.Contracts
{
    public interface IAuthService
    {
        string GenerateJwtToken(ApplicationUser user);
        string GenerateSecureCode();
        Task SaveOtpCode(string phoneNumber, string code);
    }
}
