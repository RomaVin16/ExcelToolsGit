using Microsoft.AspNetCore.Mvc;
using API.Models;
using ExcelTools;

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
            await using var stream = request.File.OpenReadStream();
            
            var fileId = fileProcessor.Upload(request.File.FileName, stream);

            return Ok(fileId);
        }

        [HttpGet("download/{fileId}")]
        public async Task<IActionResult> DownloadFile(Guid fileId)
        {
            var fileStream = fileProcessor.Download(fileId);

            var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

                return File(fileStream.FileStream, contentType, fileStream.FileName);
            }
    }
}
