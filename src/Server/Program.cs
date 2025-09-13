using DDS.API.Data;
using Microsoft.EntityFrameworkCore;
using Serilog;
using Server.Infrastructure.Logging;
using Server.Repositories;
using Server.Services;
using Server.Shared;
using Server.Shared.BaiFileContext;
using System.Text.Json;

internal class Program
{
    private static void Main(string[] args)
    
    {
;

        var builder = WebApplication.CreateBuilder(args);

        Log.Logger = new LoggerConfiguration()
            .WriteTo.File("Logs/dds-log-.txt", rollingInterval: RollingInterval.Day)
            .Enrich.FromLogContext()
            .CreateLogger();

        builder.Services.AddControllers()
     .AddJsonOptions(options =>
     {
         options.JsonSerializerOptions.PropertyNamingPolicy = JsonNamingPolicy.CamelCase;
     });


        builder.Host.UseSerilog();
        builder.Services.AddScoped(typeof(IAppLogger<>), typeof(AppLogger<>));
         
        builder.Services.AddEndpointsApiExplorer();
        builder.Services.AddSwaggerGen();
        builder.Services.AddControllers();
        builder.Services.AddScoped < ILoadBankService, LoadBankService>();
        builder.Services.AddScoped<ILoadBankRepository, LoadBankRepository>();
        builder.Services.AddScoped<IAuthenticationService, AuthenticationService>();
        builder.Services.AddScoped<IValidationService, ValidationService>();
        builder.Services.AddScoped<ITransactionService, TransactionsService>();
        builder.Services.AddScoped<IInvalidRecordRepository, InvalidRecordRepository>();
        builder.Services.AddScoped<IPreEditService, PreEditService>();
        builder.Services.AddScoped<IncomeSourceTypeRepository>();
        builder.Services.AddScoped<IIncomeSourceTypeService, IncomeSourceTypeService>();
        builder.Services.AddSingleton<DapperDbContext>();
        builder.Services.AddScoped<IInstitutionRepository, InstitutionRepository>();
        builder.Services.AddScoped<IInstitutionService, InstitutionService>();
        builder.Services.AddScoped<IDDConfigRepository , DDConfigRepository>();
        builder.Services.AddScoped<IDDConfigService, DDConfigService>();
        builder.Services.AddScoped<IRegionRepository, RegionRepository>();
        builder.Services.AddScoped<IRegionService, RegionService>();
        builder.Services.AddScoped<IOverrideService, OverrideService>();
        builder.Services.AddScoped<IPostDepositsService, PostDepositsService>();
        builder.Services.AddScoped<IPostDepositsRepository, PostDepositsRepository>();
        builder.Services.AddScoped<ILookups, LookupsRepository>();
        builder.Services.AddScoped<ILookupService, LookupService>();
        builder.Services.AddScoped<IBalanceService, BalanceService>();
        builder.Services.AddScoped<IBalanceRepository, BalanceRepository>();

        builder.Services.AddSingleton<IBaiFileContext, BaiFileContext>();
        builder.Services.AddCors(options =>
        {
            options.AddPolicy("AllowAll",
                policy => policy.AllowAnyOrigin().AllowAnyMethod().AllowAnyHeader());
        });


        builder.Services.Configure<PostingConfig>(builder.Configuration.GetSection("PostingConfig"));

        var app = builder.Build();
        app.MapControllers();
        if (app.Environment.IsDevelopment())
        {
            app.UseSwagger();
            app.UseSwaggerUI();
        }
        app.UseCors("AllowAll");
        app.UseAuthorization();
        app.UseHttpsRedirection();

        app.Run();
    }
}