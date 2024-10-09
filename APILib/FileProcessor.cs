
using API.Models;
using APILib;
using APILib.Contracts;
using Microsoft.Extensions.Configuration;

namespace ExcelTools
{
    public class FileProcessor: IFileProcessor
    {
        private readonly FileService service;
        private readonly FileRepository fileRepository;
        private readonly IArchiveService archiveService;

        public FileProcessor(IConfiguration configuration)
        {
            fileRepository = new FileRepository(configuration);
            service = new FileService(configuration, fileRepository, archiveService);
        }

        public async Task<Guid> UploadAsync(string fileName, Stream stream)
        {
            var (fileId, folderId) = service.CreateFolder();

            var filePath = Path.Combine(folderId, fileName);

            await using var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write);

            await stream.CopyToAsync(fileStream);

            await fileRepository.Create(stream, fileId, fileName);

            return fileId;
        }

        public FileResult Download(Guid fileId)
        {
            return service.Get(fileId);
        }
    }
}