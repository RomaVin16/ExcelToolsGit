using System.IO.Compression;
using APILib.Contracts;

namespace APILib
{
    public class ArchiveService : IArchiveService
    {
        /// <summary>
        /// Архивирование файлов 
        /// </summary>
        /// <param name="tempZipPath"></param>
        /// <param name="tempFolderPath"></param>
        public void ArchiveFile(string tempZipPath, string tempFolderPath)
        {
            using var zipStream = new FileStream(tempZipPath, FileMode.Create, FileAccess.Write, FileShare.None);
            using (var zip = new ZipArchive(zipStream, ZipArchiveMode.Create))
            {
                foreach (var file in Directory.GetFiles(tempFolderPath))
                {
                    zip.CreateEntryFromFile(file, Path.GetFileName(file));
                }
            }
        }
    }
}
