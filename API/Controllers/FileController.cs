using APILib;
using Microsoft.AspNetCore.Mvc;
using System.ComponentModel.DataAnnotations;

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

                var fileName = $"{fileId}.xlsx";
                var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

                return File(fileStream, contentType, fileName);
            }
    }

    public class FileUploadRequest
    {
        [Required]
        [FromForm]
        public IFormFile File { get; set; }
    }
}
