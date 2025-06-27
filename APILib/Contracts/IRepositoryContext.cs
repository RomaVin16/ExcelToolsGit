using API.Models;
using APILib.Repository;
using System.Threading.Tasks;

namespace APILib.Contracts
{
    public interface IRepositoryContext
    {
        Task Create(Stream stream, Guid fileId, string fileName, int operation);
        string GetFileName(Guid fileId);


        Task<OtpCode?> GetValidOtpCodeAsync(string phoneNumber, string code);
        Task CreateOtpCodeAsync(OtpCode otp);

        Task<int> DeleteAllForPhoneAsync(string phoneNumber);

        Task CreateHistory(string userId, Guid fileId, string operation);

		Task<List<ProcessedFileHistory>>  GetHistory(string userId);

    }
}
