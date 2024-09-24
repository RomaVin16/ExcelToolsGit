using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
            var folderPath = Path.Combine(rootPath, subfolder1, subfolder2, folderName).Replace(Path.DirectorySeparatorChar, '/');

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
var folderPath = Path.Combine(rootPath, subfolder1, subfolder2, folderName).Replace(Path.DirectorySeparatorChar, '/');

return folderPath;
        }

        public FileResult Get(Guid fileId)
        {
            using var db = new FileRepository();
            var result = new FileResult();

                var fileName = db.Files.Where(y => y.Id == fileId)
                    .Select(x => x.FileName)
                    .FirstOrDefault();

                var filePath = GetFolder(fileId);

                var stream = File.OpenRead(Path.Combine(filePath, fileName).Replace('/', Path.DirectorySeparatorChar));

                result.FileName = fileName;
                result.FileStream = stream;

                return result;
        }
    }
}
