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

        public FileController(IFileProcessor fileProcessor)
        {
            this.fileProcessor = fileProcessor;
        }

        [HttpPost("upload")]
        public async Task<IActionResult> UploadFile([FromForm] FileUploadRequest files)
        {
            var fileIds = new List<Guid>();

			if (files.Files.Length == 1)
            {
                await using var stream = files.Files[0].OpenReadStream();

                var fileId = await fileProcessor.UploadAsync(files.Files[0].FileName, stream);
                fileIds.Add(fileId);
            }
            else
            {
                foreach (var file in files.Files)
                {
                    await using var stream = file.OpenReadStream();

                    var fileId = await fileProcessor.UploadAsync(file.FileName, stream);
                    fileIds.Add(fileId);
                }
            }

            return Ok(fileIds);
        }

        [HttpGet("download/{fileId}")]
        public Task<IActionResult> DownloadFile(Guid fileId)
        {
            var fileStream = fileProcessor.Download(fileId);


            //var result = fileService.Get(fileId);

            //var contentType = fileService.GetMime(result.FileName);
            const string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

            return Task.FromResult<IActionResult>(File(fileStream.FileStream, contentType, fileStream.FileName));
        }
    }
}
