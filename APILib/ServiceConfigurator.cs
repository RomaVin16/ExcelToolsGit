using APILib.Contracts;
using APILib.Repository;
using APILib.Services;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace APILib
{
    public static class ServiceConfigurator
    {
        public static void ConfigureServices(IServiceCollection services, IConfiguration configuration)
        {
            services.AddDbContext<RepositoryContext>(options =>
                options.UseSqlServer(configuration.GetConnectionString("DefaultConnection")));

            services.AddScoped<IRepositoryContext, RepositoryContext>();
            services.AddScoped<IFileService, FileService>();
            services.AddScoped<IFileProcessor, FileProcessor>();
            services.AddScoped<IArchiveService, ArchiveService>();
            services.AddScoped<IAuthService, AuthService>();
            services.AddSingleton<ISmsService, SmsRuService>();
        }
    }
}
