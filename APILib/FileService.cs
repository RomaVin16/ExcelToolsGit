using Microsoft.Extensions.Configuration;
using ExcelTools.Cleaner;
using ExcelTools.ColumnSplitter;
using ExcelTools.DuplicateRemover;
using ExcelTools.Merger;
using ExcelTools.Splitter;
using Microsoft.AspNetCore.StaticFiles;
using APILib.APIOptions;
using FileResult = API.Models.FileResult;
using APILib.Contracts;
using ExcelTools.Rotate;


namespace APILib
{
    public class FileService: IFileService
    {
        private readonly string rootPath;
        private readonly IFileRepository _fileRepository;
        private readonly IArchiveService _archiveService;

        public FileService(IConfiguration configuration, IFileRepository fileRepository, IArchiveService archiveService)
        {
            rootPath = configuration["FileStorage:RootPath"];
            _fileRepository = fileRepository;
            _archiveService = archiveService;
        }

        /// <summary>
        /// Создание новой папки 
        /// </summary>
        /// <returns></returns>
        public (Guid fileId, string folderPath) CreateFolder()
        {
            var fileId = Guid.NewGuid();

            var folderPath = GetFolder(fileId);

            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            return (fileId, folderPath);
        }

        /// <summary>
        /// Получение пути к папке по Guid 
        /// </summary>
        /// <param name="fileId"></param>
        /// <returns></returns>
        public string GetFolder(Guid fileId)
        {
            var folderName = fileId.ToString();
            var subfolder1 = folderName.Substring(0, 2);
            var subfolder2 = folderName.Substring(2, 2);
            var folderPath = Path.Combine(rootPath, subfolder1, subfolder2, folderName);

            return folderPath;
        }

        /// <summary>
        /// Получение информации о файле 
        /// </summary>
        /// <param name="fileId"></param>
        /// <param name="fileRepository"></param>
        /// <returns></returns>
        public FileResult Get(Guid fileId)
        {
            var result = new FileResult();

            var fileName = _fileRepository.GetFileName(fileId);

            var filePath = GetFolder(fileId);

            var stream = File.OpenRead(Path.Combine(filePath, fileName));

            result.FileName = fileName;
            result.FileStream = stream;

            return result;
        }

        public string DetermineSavePath(int fileNumber)
        {
            var (resultFileId, resultFolderId) = CreateFolder();

 return Path.Combine(resultFolderId, $"file_part{fileNumber}.xlsx");
        }

        public Guid Clean(CleanerAPIOptions options)
        {
            var cleaner = new Cleaner();
            var fileResult = new FileResult
            {
                FileName = _fileRepository.GetFileName(options.FileId)
            };

            var (resultFileId, resultFolderId) = CreateFolder();

            var result = cleaner.Process(new CleanOptions
            {
                FilePath = Path.Combine(GetFolder(options.FileId), fileResult.FileName),
                ResultFilePath = Path.Combine(resultFolderId, fileResult.FileName),
                SheetNumber = options.SheetNumber,
                SkipRows = options.SkipRows
            });

            var filePath = Path.Combine(resultFolderId, fileResult.FileName);
            var stream = File.OpenRead(filePath);

            _fileRepository.Create(stream, resultFileId, fileResult.FileName);

            return resultFileId;
        }

        public Guid DuplicateRemove(DuplicateRemoverAPIOptions options)
        {
            var duplicateRemover = new DuplicateRemover();
            var fileResult = new FileResult();

            fileResult.FileName = _fileRepository.GetFileName(options.FileId);

            var (resultFileId, resultFolderId) = CreateFolder();

            var result = duplicateRemover.Process(new DuplicateRemoverOptions()
            {
                FilePath = Path.Combine(GetFolder(options.FileId), fileResult.FileName),
                ResultFilePath = Path.Combine(resultFolderId, fileResult.FileName),
                KeysForRowsComparison = options.KeysForRowsComparison,
                SheetNumber = options.SheetNumber,
                SkipRows = options.SkipRows
            });

            var filePath = Path.Combine(resultFolderId, fileResult.FileName);
            var stream = File.OpenRead(filePath);

            _fileRepository.Create(stream, resultFileId, fileResult.FileName);

            return resultFileId;
        }

        public Guid Merge(MergerAPIOptions options)
        {
            var merger = new Merger();
            var fileResult = new FileResult();
            var array = new string[options.MergeFilePaths.Length];

            for (var i = 0; i < options.MergeFilePaths.Length; i++)
            {
                var fileName = _fileRepository.GetFileName(options.MergeFilePaths[i]);

                array[i] = Path.Combine(GetFolder(options.MergeFilePaths[i]), fileName);
            }

            var (resultFileId, resultFolderId) = CreateFolder();

            var result = merger.Process(new MergerOptions
            {
                MergeFilePaths = array,
                ResultFilePath = Path.Combine(resultFolderId, "merge_result.xlsx"),
                SheetNumber = options.SheetNumber,
                SkipRows = options.SkipRows,
                MergeMode = (MergerOptions.MergeType)options.MergeMode
            });

            var filePath = Path.Combine(resultFolderId, fileResult.FileName);
            var stream = File.OpenRead(filePath);

            _fileRepository.Create(stream, resultFileId, fileResult.FileName);

            return resultFileId;
        }

        public Guid Split(SplitterAPIOptions options)
        {
            var splitter = new Splitter();
            var fileResult = new FileResult
            {
                FileName = _fileRepository.GetFileName(options.FileId)
            };

            var baseFileName = Path.GetFileNameWithoutExtension(fileResult.FileName);
            var extension = Path.GetExtension(fileResult.FileName);

            var tempFolderPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempFolderPath);

            var splitterOptions = new SplitterOptions
            {
                FilePath = Path.Combine(GetFolder(options.FileId), fileResult.FileName),
                ResultFilePath = Path.Combine(tempFolderPath, baseFileName + "_{0}" + extension),
                //IndividualPathToEachFile = DetermineSavePath,
                SheetNumber = options.SheetNumber,
                SkipRows = options.SkipRows,
                ResultsCount = options.ResultsCount,
                SplitMode = (SplitterOptions.SplitType)options.SplitMode,
                AddHeaderRows = options.AddHeaderRows
            };

            var result = splitter.Process(splitterOptions);

            var (resultZipId, resultFolderWithZipId) = CreateFolder();

            if (splitterOptions.IndividualPathToEachFile == null)
            {
                var tempZipPath = Path.Combine(resultFolderWithZipId, $"{baseFileName}_results.zip");

                _archiveService.ArchiveFile(tempZipPath, tempFolderPath);

                var stream = File.OpenRead(tempZipPath);

                _fileRepository.Create(stream, resultZipId, $"{baseFileName}_results.zip");

                Directory.Delete(tempFolderPath, true);
            }
            else
            {
//Здесь должна быть логика обработки файлов через делегаты
            }


            return resultZipId;
        }

        public Guid SplitColumn(ColumnSplitterAPIOptions options)
        {
            var columnSplitter = new ColumnSplitter();
            var fileResult = new FileResult
            {
                FileName = _fileRepository.GetFileName(options.FileId),
            };

            var (resultFileId, resultFolderId) = CreateFolder();

            var result = columnSplitter.Process(new ColumnSplitterOptions
            {
FilePath = Path.Combine(GetFolder(options.FileId), fileResult.FileName),
ResultFilePath = Path.Combine(resultFolderId, fileResult.FileName),
                SplitSymbols = options.SplitSymbols,
                ColumnName = options.ColumnName,
                SkipHeaderRows = options.SkipHeaderRows,
                SkipRows = options.SkipHeaderRows
            });

            var filePath = Path.Combine(resultFolderId, fileResult.FileName);
            var stream = File.OpenRead(filePath);

            _fileRepository.Create(stream, resultFileId, fileResult.FileName);

            return resultFileId;
        }

        public Guid Rotate(RotaterAPIOptions options)
        {
            var rotater = new Rotater();
            var fileResult = new FileResult
            {
                FileName = _fileRepository.GetFileName(options.FileId)
            };

            var (resultFileId, resultFolderId) = CreateFolder();

            var result = rotater.Process(new RotaterOptions
            {
                FilePath = Path.Combine(GetFolder(options.FileId), fileResult.FileName),
                ResultFilePath = Path.Combine(resultFolderId, fileResult.FileName),
                SheetNumber = options.SheetNumber,
                SkipRows = options.SkipRows
            });

            var filePath = Path.Combine(resultFolderId, fileResult.FileName);
            var stream = File.OpenRead(filePath);

            _fileRepository.Create(stream, resultFileId, fileResult.FileName);

            return resultFileId;
        }

        public string GetMime(string fileName)
        {
            var provider = new FileExtensionContentTypeProvider();

            if (!provider.TryGetContentType(fileName, out var contentType))
            {
                contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            }

            return contentType;
        }
    }
}
    