using APILib;
using APILib.APIOptions;
using APILib.Contracts;
using ExcelTools;
using Microsoft.AspNetCore.Mvc;

namespace API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        private readonly IFileService fileService;

        public ExcelController(IFileService fileService)
        {
            this.fileService = fileService;
        }

        [HttpPost("clean")]
        public async Task<IActionResult> Clean([FromBody] CleanerAPIOptions options)
        {
            var resultFileId = fileService.Clean(options);

            return Ok(resultFileId);
        }

        [HttpPost("duplicateRemove")]
        public async Task<IActionResult> DuplicateRemove([FromBody] DuplicateRemoverAPIOptions options)
        {
            var resultFileId = fileService.DuplicateRemove(options);

            return Ok(resultFileId);
        }

        [HttpPost("merge")]
        public async Task<IActionResult> Merge([FromBody] MergerAPIOptions options)
        {
            var resultFileId = fileService.Merge(options);

            return Ok(resultFileId);
        }

        [HttpPost("split")]
        public async Task<IActionResult> Split([FromBody] SplitterAPIOptions options)
        {
            var resultFileId = fileService.Split(options);

            return Ok(resultFileId);
        }

        [HttpPost("columnSplit")]
        public async Task<IActionResult> SplitColumn([FromBody] ColumnSplitterAPIOptions options)
        {
            var resultFileId = fileService.SplitColumn(options);

            return Ok(resultFileId);
        }

        [HttpPost("rotate")]
        public async Task<IActionResult> Rotate([FromBody] RotaterAPIOptions options)
        {
            var resultFileId = fileService.Rotate(options);

            return Ok(resultFileId);
        }
    }
}
