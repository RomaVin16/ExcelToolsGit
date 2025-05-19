using System.Security.Claims;
using APILib.APIOptions;
using APILib.Contracts;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        private readonly IFileService _fileService;
        private readonly IRepositoryContext _repositoryContext;
		private string? _userId;


		public ExcelController(IFileService fileService, IRepositoryContext repositoryContext)
        {
            _fileService = fileService;
            _repositoryContext = repositoryContext;
            }

        [Authorize]
        [HttpPost("clean")]
        public async Task<IActionResult> Clean([FromBody] CleanerAPIOptions options)
        {
	        var _userId = User.FindFirstValue(ClaimTypes.NameIdentifier);

			var resultFileId = _fileService.Clean(options);

			await _repositoryContext.CreateHistory(_userId, resultFileId, "Clean");

			return Ok(resultFileId);
        }

        [HttpPost("duplicateRemove")]
        public async Task<IActionResult> DuplicateRemove([FromBody] DuplicateRemoverAPIOptions options)
        {
            var resultFileId = _fileService.DuplicateRemove(options);

            await _repositoryContext.CreateHistory(_userId, resultFileId, "DuplicateRemove");

			return Ok(resultFileId);
        }

        [HttpPost("merge")]
        public async Task<IActionResult> Merge([FromBody] MergerAPIOptions options)
        {
            var resultFileId = _fileService.Merge(options);

			await _repositoryContext.CreateHistory(_userId, resultFileId, "Merge");

			return Ok(resultFileId);
        }

        [HttpPost("split")]
        public async Task<IActionResult> Split([FromBody] SplitterAPIOptions options)
        {
            var resultFileId = _fileService.Split(options);

            await _repositoryContext.CreateHistory(_userId, resultFileId, "Split");

			return Ok(resultFileId);
        }

        [HttpPost("columnSplit")]
        public async Task<IActionResult> SplitColumn([FromBody] ColumnSplitterAPIOptions options)
        {
            var resultFileId = _fileService.SplitColumn(options);

            await _repositoryContext.CreateHistory(_userId, resultFileId, "SplitColumn");

			return Ok(resultFileId);
        }

        [HttpPost("rotate")]
        public async Task<IActionResult> Rotate([FromBody] RotaterAPIOptions options)
        {
            var resultFileId = _fileService.Rotate(options);

            await _repositoryContext.CreateHistory(_userId, resultFileId, "Rotate");

			return Ok(resultFileId);
        }
    }
}
