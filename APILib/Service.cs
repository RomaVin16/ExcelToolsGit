using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;

namespace APILib
{
    public class Service
    {
        private readonly string rootPath;

        public Service(IConfiguration configuration)
        {
            rootPath = configuration["FileStorage:RootPath"];
        }

        public (Guid fileId, string folderPath) CreateFolder()
        {
            Guid fileId = Guid.NewGuid();
            var folderName = fileId.ToString();

            var subfolder1 = folderName.Substring(0, 2);
            var subfolder2 = folderName.Substring(2, 2);
            var folderPath = Path.Combine(rootPath, subfolder1, subfolder2, folderName).Replace("\\", "/");

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
var folderPath = Path.Combine(rootPath, subfolder1, subfolder2, folderName).Replace("\\", "/"); ;

return folderPath;
        }
    }
}
