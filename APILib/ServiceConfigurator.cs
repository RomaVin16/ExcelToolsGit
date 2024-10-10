using APILib.Contracts;
using ExcelTools;
using Microsoft.Extensions.DependencyInjection;

namespace APILib
{
    public static class ServiceConfigurator
    {
        public static void ConfigureServices(IServiceCollection services)
        {
            services.AddScoped<IFileRepository, FileRepository>();
            services.AddScoped<IFileService, FileService>();
            services.AddScoped<IFileProcessor, FileProcessor>();
            services.AddScoped<IArchiveService, ArchiveService>();
        }
    }
}
