using System;
using System.Data;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace DDS.API.Data
{
    public class DapperDbContext
    {
        private readonly IConfiguration _configuration;
        private readonly string _ddsConnectionString;
        private readonly string _pfsConnectionString;

        public DapperDbContext(IConfiguration configuration)
        {
            _configuration = configuration;
            _ddsConnectionString = _configuration.GetConnectionString("DDSConnection")
                ?? throw new ArgumentNullException("DDSConnection is missing in appsettings.json");
            _pfsConnectionString = _configuration.GetConnectionString("PFSConnection")
                ?? throw new ArgumentNullException("PFSConnection is missing in appsettings.json");
        }

        public IDbConnection CreateDDSConnection() => new SqlConnection(_ddsConnectionString);
        public IDbConnection CreatePFSConnection() => new SqlConnection(_pfsConnectionString);

        public IDbConnection CreateConnection(string schema)
        {
            return schema.ToLower() switch
            {
                "dds_schema" => new SqlConnection(_ddsConnectionString),
                "pfs_schema" => new SqlConnection(_pfsConnectionString),
                _ => throw new ArgumentException($"Invalid schema name: {schema}")
            };
        }
    }
}
