using API.Models;
using Microsoft.Extensions.Configuration;

namespace APILib
{
    public class FileService
    {
        private readonly string rootPath;

        public FileService(IConfiguration configuration)
        {
            rootPath = configuration["FileStorage:RootPath"];
        }

        public (Guid fileId, string folderPath) CreateFolder()
        {
            Guid fileId = Guid.NewGuid();
            var folderName = fileId.ToString();

            var subfolder1 = folderName.Substring(0, 2);
            var subfolder2 = folderName.Substring(2, 2);
            var folderPath = Path.Combine(rootPath, subfolder1, subfolder2, folderName);

            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            return (fileId, folderPath);
        }

        public string GetFolder(Guid fileId)
        {
var folderName = fileId.ToString();
var subfolder1 = folderName.Substring(0, 2);
var subfolder2 = folderName.Substring(2, 2);
var folderPath = Path.Combine(rootPath, subfolder1, subfolder2, folderName);

return folderPath;
        }

        public FileResult Get(Guid fileId, FileRepository db)
        {
            var result = new FileResult();

                var fileName = db.Files.Where(y => y.Id == fileId)
                    .Select(x => x.FileName)
                    .FirstOrDefault();

                var filePath = GetFolder(fileId);

            var stream = File.OpenRead(Path.Combine(filePath, fileName));

            result.FileName = fileName;
                result.FileStream = stream;

                return result;
        }
    }
}
