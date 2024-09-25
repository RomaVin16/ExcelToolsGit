using API.Models;
using Microsoft.Extensions.Configuration;
using ExcelTools.Abstraction;
using ExcelTools.Cleaner;
using ExcelTools.ColumnSplitter;
using ExcelTools.DuplicateRemover;
using ExcelTools.Merger;
using ExcelTools.Splitter;
using System.Reflection;

namespace APILib
{
    public class FileService
    {
        private readonly string rootPath;
        private readonly FileRepository _db;

        public FileService(IConfiguration configuration, FileRepository db)
        {
            rootPath = configuration["FileStorage:RootPath"];
            _db = db;
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

        public Guid Clean(Guid fileId)
        {
            var cleaner = new Cleaner();

            var fileName = _db.Files.Where(y => y.Id == fileId)
                .Select(x => x.FileName)
                .FirstOrDefault();

            var (resultFileId, resultFolderId) = CreateFolder();

            var result = cleaner.Process(new CleanOptions
            {
                FilePath = Path.Combine(GetFolder(fileId), fileName),
                ResultFilePath = Path.Combine(resultFolderId, fileName)
            });

            var filePath = Path.Combine(resultFolderId, fileName);
            var stream = File.OpenRead(filePath);

            _db.Create(_db, stream, resultFileId, fileName);

            return resultFileId;
        }
    }
}
