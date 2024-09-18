using Microsoft.Extensions.Configuration;

namespace APILib
{
    public class FileProcessor
    {
        private readonly Service service;

        public FileProcessor(IConfiguration configuration)
        {
            service = new Service(configuration);
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
            try
            {
                var filePath = Path.Combine(service.GetFolder(fileId));

                var fileData = File.ReadAllBytes(Path.Combine(filePath, $"{fileId}.xlsx"));
                return new MemoryStream(fileData);
            }
            catch (Exception ex)
            {
Console.WriteLine("Ошибка: " + ex.Message);
            }

            return null;
        }
    }
}
