using Client;
using Client.Library;
using Client.Services;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
using Radzen;
//using Blazored.LocalStorage;


var builder = WebAssemblyHostBuilder.CreateDefault(args);
builder.RootComponents.Add<App>("#app");
builder.RootComponents.Add<HeadOutlet>("head::after");
//builder.Services.AddBlazoredLocalStorage();

builder.Services.AddScoped(sp => new HttpClient { BaseAddress = new Uri(builder.HostEnvironment.BaseAddress) });

//Radzen Services
builder.Services.AddScoped<DialogService>();
builder.Services.AddScoped<NotificationService>();
builder.Services.AddScoped<TooltipService>();
builder.Services.AddScoped<ContextMenuService>();


builder.Services.AddScoped<AuthenticationService>();
builder.Services.AddScoped<FileloadService>();
builder.Services.AddScoped<ValidationService>();
builder.Services.AddSingleton<ConfigurationService>();
builder.Services.AddScoped<PreEditService>();
builder.Services.AddScoped<IncomeSourceTypeService>();
builder.Services.AddScoped<InstitutionService>();
builder.Services.AddScoped<DDConfigService>();
builder.Services.AddScoped<RegionService>();
builder.Services.AddScoped<LookupApiService>();
builder.Services.AddScoped<INewTabService, NewTabService>();



builder.Services.AddBlazorBootstrap();

await builder.Build().RunAsync();
