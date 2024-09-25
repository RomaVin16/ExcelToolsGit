
using API.Models;
using APILib;
using Microsoft.Extensions.Configuration;

namespace ExcelTools
{
    public class FileProcessor
    {
        private readonly FileService service;
        private readonly FileRepository db;

        public FileProcessor(IConfiguration configuration)
        {
            service = new FileService(configuration);
            db = new FileRepository(configuration);
        }

        public Guid Upload(string fileName, Stream stream)
        {
            var (fileId, folderId) = service.CreateFolder();

            var filePath = Path.Combine(folderId, fileName);

            using var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write);

            stream.CopyTo(fileStream);

                db.Create(db, stream, fileId, fileName);
                
            return fileId;
        }

        public FileResult Download(Guid fileId)
        {
            return service.Get(fileId, db);
        }
    }
}