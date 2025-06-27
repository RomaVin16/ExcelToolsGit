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

		public ExcelController(IFileService fileService, IRepositoryContext repositoryContext)
        {
            _fileService = fileService;
            _repositoryContext = repositoryContext;
            }

        [Authorize]
        [HttpPost("clean")]
        public async Task<IActionResult> Clean([FromBody] CleanerAPIOptions options)
        {
	        var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
			var resultFileId = _fileService.Clean(options);

			await _repositoryContext.CreateHistory(userId, resultFileId, "Clean");

			return Ok(resultFileId);
        }

        [Authorize]
        [HttpPost("removeduplicates")]
        public async Task<IActionResult> DuplicateRemove([FromBody] DuplicateRemoverAPIOptions options)
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            var resultFileId = _fileService.DuplicateRemove(options);

            await _repositoryContext.CreateHistory(userId, resultFileId, "DuplicateRemove");

			return Ok(resultFileId);
        }

        [Authorize]
        [HttpPost("merge")]
        public async Task<IActionResult> Merge([FromBody] MergerAPIOptions options)
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            var resultFileId = _fileService.Merge(options);

			await _repositoryContext.CreateHistory(userId, resultFileId, "Merge");

			return Ok(resultFileId);
        }

        [Authorize]
        [HttpPost("split")]
        public async Task<IActionResult> Split([FromBody] SplitterAPIOptions options)
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            var resultFileId = _fileService.Split(options);

            await _repositoryContext.CreateHistory(userId, resultFileId, "Split");

			return Ok(resultFileId);
        }

        [Authorize]
        [HttpPost("columnsplit")]
        public async Task<IActionResult> SplitColumn([FromBody] ColumnSplitterAPIOptions options)
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            var resultFileId = _fileService.SplitColumn(options);

            await _repositoryContext.CreateHistory(userId, resultFileId, "SplitColumn");

			return Ok(resultFileId);
        }

        [Authorize]
        [HttpPost("rotate")]
        public async Task<IActionResult> Rotate([FromBody] RotaterAPIOptions options)
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            var resultFileId = _fileService.Rotate(options);

            await _repositoryContext.CreateHistory(userId, resultFileId, "Rotate");

			return Ok(resultFileId);
        }
    }
}
