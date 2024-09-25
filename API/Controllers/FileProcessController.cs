using APILib;
using ExcelTools;
using Microsoft.AspNetCore.Mvc;

namespace API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class FileProcessController : ControllerBase
    {
        private readonly FileService fileService;
        private readonly FileProcessor fileProcessor;

        public FileProcessController(IConfiguration configuration)
        {
            fileService = new FileService(configuration, new FileRepository(configuration));
        }

        [HttpGet("clean/{fileId}")]
        public async Task<IActionResult> Clean(Guid fileId)
        {
            var resultFileId = fileService.Clean(fileId);

            return Ok(resultFileId);
        }
    }
}
