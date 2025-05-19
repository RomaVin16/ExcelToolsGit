using System.Net.Http.Headers;
using System.Security.Claims;
using Microsoft.AspNetCore.Components.Authorization;
using Blazored.LocalStorage;
using System.IdentityModel.Tokens.Jwt;

namespace ExcelToolsFrontend.Auth
{
	public class CustomAuthStateProvider : AuthenticationStateProvider
	{
		private readonly ILocalStorageService _localStorage;
		private readonly HttpClient _http;

		public CustomAuthStateProvider(ILocalStorageService localStorage, HttpClient http)
		{
			_localStorage = localStorage;
			_http = http;
		}

		public async Task MarkUserAsAuthenticated(string token)
		{
			await _localStorage.SetItemAsync("authToken", token);
			_http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
			NotifyAuthenticationStateChanged(GetAuthenticationStateAsync());
		}

		public async Task MarkUserAsLoggedOut()
		{
			await _localStorage.RemoveItemAsync("authToken");
			_http.DefaultRequestHeaders.Authorization = null;
			NotifyAuthenticationStateChanged(GetAuthenticationStateAsync());
		}

		public override async Task<AuthenticationState> GetAuthenticationStateAsync()
		{
			var token = await _localStorage.GetItemAsync<string>("authToken");

			if (string.IsNullOrEmpty(token))
			{
				return new AuthenticationState(new ClaimsPrincipal(new ClaimsIdentity()));
			}

			var identity = new ClaimsIdentity(ParseClaimsFromJwt(token), "jwt");
			return new AuthenticationState(new ClaimsPrincipal(identity));
		}

		private static IEnumerable<Claim> ParseClaimsFromJwt(string jwt)
		{
			var handler = new JwtSecurityTokenHandler();
			var token = handler.ReadJwtToken(jwt);
			return token.Claims;
		}
	}
}