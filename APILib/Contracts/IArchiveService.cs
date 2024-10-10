namespace APILib.Contracts
{
    public interface IArchiveService
    {
        void ArchiveFile(string tempZipPath, string tempFolderPath);
    }
}
