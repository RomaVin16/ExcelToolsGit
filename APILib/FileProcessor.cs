
using API.Models;
using APILib;
using APILib.Repository;
using Microsoft.Extensions.Configuration;
using static APILib.Repository.Files;

namespace ExcelTools
{
    public class FileProcessor
    {
        private readonly FileService service;

        public FileProcessor(IConfiguration configuration)
        {
            service = new FileService(configuration);
        }

        public Guid Upload(string fileName, Stream stream)
        {
            var (fileId, folderId) = service.CreateFolder();

            var filePath = Path.Combine(folderId, fileName);

            using var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write);

            stream.CopyTo(fileStream);

            using (var db = new FileRepository())
            {
                db.Create(db, stream, fileId, fileName);
            }

            return fileId;
        }

        public FileResult Download(Guid fileId)
        {
            return service.Get(fileId);
        }
    }
}