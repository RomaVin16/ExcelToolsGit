using APILib.Contracts;
using APILib.Repository;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using static APILib.Repository.Files;

namespace APILib
{
    public sealed class FileRepository: DbContext, IFileRepository
    {
        public DbSet<Files> Files => Set<Files>();
        private readonly string connectionString;

        public FileRepository(IConfiguration configuration)
        {
            connectionString = configuration["ConnectionStrings:DefaultConnection"];
        }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseNpgsql(connectionString);
        }

        public async Task Create(Stream stream, Guid fileId, string fileName)
        {
            var fileInfo = new Files
                {
                    Id = fileId,
                    FileName = fileName,
                    SaveDate = DateTime.UtcNow,
                    State = FileState.Active,
                    SizeBytes = stream.Length,
                };

                Files.Add(fileInfo);
                await SaveChangesAsync();
        }

        public string GetFileName(Guid fileId)
        {
            var fileName = Files.Where(y => y.Id == fileId)
                .Select(x => x.FileName)
                .FirstOrDefault();

            return fileName;
        }

        public void Update()
        {
        }
    }
}
