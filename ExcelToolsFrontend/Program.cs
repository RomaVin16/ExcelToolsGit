using Blazored.LocalStorage;
using ExcelToolsFrontend.Auth;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
using MudBlazor.Services;

namespace ExcelToolsFrontend
{
	public class Program
	{
		public static async Task Main(string[] args)
		{
			var builder = WebAssemblyHostBuilder.CreateDefault(args);

			builder.Services.AddScoped<CustomAuthStateProvider>();
			builder.Services.AddScoped<AuthenticationStateProvider>(sp =>
				sp.GetRequiredService<CustomAuthStateProvider>());
			builder.Services.AddAuthorizationCore();
			builder.Services.AddBlazoredLocalStorage();
			builder.Services.AddMudServices();

			builder.RootComponents.Add<App>("#app");
			builder.RootComponents.Add<HeadOutlet>("head::after");

			builder.Services.AddScoped(sp => new HttpClient
			{
				BaseAddress = new Uri("https://localhost:5001")
			});

			await builder.Build().RunAsync();
		}


	}
}
