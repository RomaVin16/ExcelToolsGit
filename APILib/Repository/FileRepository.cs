using APILib.Repository;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using static APILib.Repository.Files;

namespace APILib
{
    public sealed class FileRepository: DbContext
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

        public Task Create(FileRepository db, Stream stream, Guid fileId, string fileName)
        {
            var fileInfo = new Files
                {
                    Id = fileId,
                    FileName = fileName,
                    SaveDate = DateTime.UtcNow,
                    State = FileState.Active,
                    SizeBytes = stream.Length,
                };

                db.Add(fileInfo);
                db.SaveChangesAsync();
                return Task.CompletedTask;
        }

        public void Update()
        {
        }
    }
}
