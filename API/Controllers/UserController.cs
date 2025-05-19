using APILib.Contracts;
using Microsoft.AspNetCore.Mvc;
using System.Security.Claims;

namespace API.Controllers
{
	[Route("api/[controller]")]
	[ApiController]
	public class UserController : Controller
	{
		private readonly IRepositoryContext _repositoryContext;

		public UserController(IRepositoryContext repositoryContext)
		{
			_repositoryContext = repositoryContext;
		}

		/// <summary>
		/// Endpoint для получения истории
		/// </summary>
		/// <param name="phoneNumber"></param>
		/// <returns></returns>
		/// <summary>
		[HttpGet("history")]
		public async Task<IActionResult> GetHistory()
		{
			var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);

			if (string.IsNullOrEmpty(userId))
			{
				return Unauthorized("User ID not found in claims.");
			}

			var history = await _repositoryContext.GetHistory(userId);
			return Ok(history);
		}
	}
}
