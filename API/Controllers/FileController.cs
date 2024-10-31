using Microsoft.AspNetCore.Mvc;
using API.Models;
using APILib.Contracts;

namespace API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class FileController : ControllerBase
    {
        private readonly IFileProcessor fileProcessor;
        private readonly IFileService fileService;

        public FileController(IFileProcessor fileProcessor, IFileService fileService)
        {
            this.fileProcessor = fileProcessor;
            this.fileService = fileService;
        }

        [HttpPost("upload")]
        public async Task<IActionResult> UploadFile([FromForm] FileUploadRequest request)
        {
            var fileIds = new List<Guid>();

            if (request.Files.Length == 1)
            {
                await using var stream = request.Files[0].OpenReadStream();

                var fileId = await fileProcessor.UploadAsync(request.Files[0].FileName, stream);
fileIds.Add(fileId);
            }
            else
            {
                foreach (var file in request.Files)
                {
                    await using var stream = file.OpenReadStream();

                    var fileId = await fileProcessor.UploadAsync(file.FileName, stream);
                    fileIds.Add(fileId);
                }
            }

            return Ok(fileIds);
        }

        [HttpGet("download/{fileId}")]
        public async Task<IActionResult> DownloadFile(Guid fileId)
        {
            var fileStream = fileProcessor.Download(fileId);

            var result = fileService.Get(fileId);

            var contentType = fileService.GetMime(result.FileName);

            return File(fileStream.FileStream, contentType, fileStream.FileName);
            }
    }
}
