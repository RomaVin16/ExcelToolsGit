using Microsoft.AspNetCore.Mvc;
using API.Models;
using ExcelTools;
using Microsoft.AspNetCore.StaticFiles;

namespace API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class FileController : ControllerBase
    {
        private readonly FileProcessor fileProcessor;


        public FileController(IConfiguration configuration)
        {
            fileProcessor = new FileProcessor(configuration);
        }

        [HttpPost("upload")]
        public async Task<IActionResult> UploadFile([FromForm] FileUploadRequest request)
        {
            using var stream = request.File.OpenReadStream();
            
            var fileId = fileProcessor.Upload(request.File.FileName, stream);

            return Ok(fileId);
        }

        [HttpGet("download/{fileId}")]
        public async Task<IActionResult> DownloadFile(Guid fileId)
        {
            var fileStream = fileProcessor.Download(fileId);

            var provider = new FileExtensionContentTypeProvider();
            string contentType;

            if (!provider.TryGetContentType(fileStream.FileName, out contentType))
            {
                contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            }

            return File(fileStream.FileStream, contentType, fileStream.FileName);
            }
    }
}
