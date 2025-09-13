using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace Server.Infrastructure.Data;

public class DapperDbContext
{
    private readonly IConfiguration _configuration;

    public DapperDbContext(IConfiguration configuration)
    {
        _configuration = configuration;
    }

    public SqlConnection CreateConnection(string schema = "dds_schema")
    {
        var connectionString = _configuration.GetConnectionString(schema) ?? 
                              _configuration.GetConnectionString("DefaultConnection");
        
        if (string.IsNullOrEmpty(connectionString))
        {
            throw new InvalidOperationException($"Connection string '{schema}' not found in configuration.");
        }

        return new SqlConnection(connectionString);
    }
} 