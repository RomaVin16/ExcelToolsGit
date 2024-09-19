using APILib;
using Microsoft.Extensions.Configuration;

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

            var filePath = Path.Combine(folderId, $"{fileId}.xlsx");

            using var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write);

            stream.CopyTo(fileStream);

            return fileId;
        }


        public Stream Download(Guid fileId)
        {
                var filePath = service.GetFolder(fileId);

                return File.OpenRead(Path.Combine(filePath, $"{fileId}.xlsx"));
        }
    }
}
