using APILib.Repository;
using Microsoft.EntityFrameworkCore;
using static APILib.Repository.Files;

namespace APILib
{
    public sealed class FileRepository: DbContext
    {
        public DbSet<Files> Files => Set<Files>();

        public FileRepository()
        {
            Database.EnsureCreated();
        }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseNpgsql("Host=localhost;Port=5432;Database=Files;Username=postgres;Password=1234");
        }

        public async Task Create(FileRepository db, Stream stream, Guid fileId, string fileName)
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
                db.SaveChanges();
        }

        public void Update()
        {
        }
    }
}
