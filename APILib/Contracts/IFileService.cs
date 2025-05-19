using APILib.APIOptions;
using API.Models;

namespace APILib.Contracts
{
    public interface IFileService
    {
        (Guid fileId, string folderPath) CreateFolder();
        string GetFolder(Guid fileId);
        FileResult Get(Guid fileId);
        string DetermineSavePath(int fileNumber);
        Guid Clean(CleanerAPIOptions options);
        Guid DuplicateRemove(DuplicateRemoverAPIOptions options);
        Guid Merge(MergerAPIOptions options);
        Guid Split(SplitterAPIOptions options);
        Guid SplitColumn(ColumnSplitterAPIOptions options);
        Guid Rotate(RotaterAPIOptions options);
        //string GetMime(string fileName);
    }
}
