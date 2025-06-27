using APILib.Contracts;
using Microsoft.AspNetCore.Identity.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using static APILib.Repository.File;

namespace APILib.Repository
{
    public sealed class RepositoryContext: IdentityDbContext<ApplicationUser>, IRepositoryContext
    {
        public DbSet<File> Files => Set<File>();

public override DbSet<ApplicationUser> Users => Set<ApplicationUser>();

public DbSet<OtpCode> OtpCodes { get; set; }

public DbSet<ProcessedFileHistory> ProcessedFileHistory{ get; set; }

		private readonly string? _connectionString;

        public RepositoryContext(IConfiguration configuration)
        {
			_connectionString = configuration["ConnectionStrings:DefaultConnection"];
		}

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseNpgsql(_connectionString);
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {

            base.OnModelCreating(modelBuilder);

            modelBuilder.Entity<ApplicationUser>()
                .Property(u => u.Id)
                .HasConversion(
                    v => Guid.Parse(v),   
                    v => v.ToString());

            modelBuilder.Entity<ApplicationUser>()
                .ToTable("aspnetusers");  


        }

        // ===== Работа с файлами =====

        /// <summary>
        /// Создание записи в табоице Files
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="fileId"></param>
        /// <param name="fileName"></param>
        /// <param name="operation"></param>
        /// <returns></returns>
        public async Task Create(Stream stream, Guid fileId, string fileName, int operation)
        {
            var fileInfo = new File
                {
                    Id = fileId,
                    FileName = fileName,
                    SaveDate = DateTime.UtcNow,
                    State = FileState.Active,
                    SizeBytes = stream.Length,
                    OperationId = operation,
			};

                Files.Add(fileInfo);
                await SaveChangesAsync();
        }

        /// <summary>
        /// Создание записи в табоице с историей
        /// </summary>
        /// <param name="userId"></param>
        /// <param name="fileId"></param>
/// <param name="operation"></param>
        /// <returns></returns>
        public async Task CreateHistory(string userId, Guid fileId, string operation)
        {
			var fileInfo = new ProcessedFileHistory
	        {
				UserId = userId,
				Operation = operation,
				FileId = fileId,
				ProcessedAt = DateTime.UtcNow,
			};

	        ProcessedFileHistory.Add(fileInfo);

	        await SaveChangesAsync();
        }

        //      /// <summary>
        ///// Получение истории 
        ///// </summary>
        ///// <param name="userId"></param>
        ///// <param name="fileId"></param>
        ///// <param name="operation"></param>
        ///// <returns></returns>
        public async Task<List<ProcessedFileHistory>> GetHistory(string userId)
        {
            return await ProcessedFileHistory
				.Where(x => x.UserId == userId)
				.OrderBy(x => x.ProcessedAt)
				.ToListAsync();
		}

        /// <summary>
        /// Получение имени файла
        /// </summary>
        /// <param name="fileId"></param>
        /// <returns></returns>
        public string GetFileName(Guid fileId)
        {
            var fileName = Files.Where(y => y.Id == fileId)
                .Select(x => x.FileName)
                .FirstOrDefault();

            return fileName;
        }

        // ===== Работа с OTP-кодами =====

        /// <summary>
        /// Сохранение OTP-кода
        /// </summary>
        public async Task SaveOtpCodeAsync(OtpCode otp)
        {
            OtpCodes.Add(otp);
            await SaveChangesAsync();
        }

        /// <summary>
        /// Получение валидного OTP-кода
        /// </summary>
        public async Task<OtpCode?> GetValidOtpCodeAsync(string phoneNumber, string code)
        {
            return await OtpCodes
                .Where(x => x.PhoneNumber == phoneNumber && x.Code == code)
                .OrderByDescending(x => x.ExpiryTime)
                .FirstOrDefaultAsync();
        }


        public async Task<int> DeleteAllForPhoneAsync(string phoneNumber)
        {
            return await OtpCodes
                .Where(x => x.PhoneNumber == phoneNumber)
                .ExecuteDeleteAsync();
        }

        public async Task CreateOtpCodeAsync(OtpCode otp)
        {
            await OtpCodes.AddAsync(otp);
            await SaveChangesAsync();
        }
    }
}

