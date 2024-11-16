using API.Models;

namespace APILib.Contracts
{
    public interface IFileProcessor
    {
        Task<Guid> UploadAsync(string fileName, Stream stream);
        FileResult Download(Guid fileId);
    }
}
