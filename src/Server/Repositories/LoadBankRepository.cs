using Dapper;
using DDS.API.Data;
using Server.Models;
using System.Data;

namespace Server.Repositories
{
	public interface ILoadBankRepository
	{
		Task<Config?> GetConfigAsync();
		Task<(int ProcResult, string BaiFileId)> InsertBaiFileSummaryAsync(DateTime baiFileDate, int fileIdNum, string createdBy);
		Task UpdateBaiFileSummaryAsync(long fileIdNum, string baiFileId, decimal availableBal, decimal collectedBal, decimal totalCredits, decimal totalDebits, decimal ledgerBal, DateTime baiFileDate);
		Task InsertWorkFileAsync(WorkFileInsert data);
		Task<bool> AnyWorkFileRecordsAsync();
	}

	public class LoadBankRepository : ILoadBankRepository
	{
		private readonly DapperDbContext _dbContext;
		private readonly ILogger<LoadBankRepository> _logger;

		public LoadBankRepository(DapperDbContext dbContext, ILogger<LoadBankRepository> logger)
		{
			_dbContext = dbContext;
			_logger = logger;
		}

		public async Task<Config?> GetConfigAsync()
		{
			using var connection = _dbContext.CreateConnection("dds_schema");
			const string sql = "SELECT SENDER_ID, RECEIVER_ID FROM DD_CONFIG_INFO";
			_logger.LogInformation("Fetching config (SENDER_ID, RECEIVER_ID) from DD_CONFIG_INFO");
			return await connection.QueryFirstOrDefaultAsync<Config>(sql);
		}

		public async Task<(int ProcResult, string BaiFileId)> InsertBaiFileSummaryAsync(DateTime baiFileDate, int fileIdNum, string createdBy)
		{
			using var connection = _dbContext.CreateConnection("dds_schema");
			var parameters = new DynamicParameters();
			parameters.Add("@bai_file_datetime", baiFileDate.ToString("yyyy-MM-dd"));
			parameters.Add("@file_id_num", fileIdNum, DbType.Int64);
			parameters.Add("@created_by", createdBy);
			parameters.Add("@available_bal", 0m, DbType.Decimal);
			parameters.Add("@collected_bal", 0m, DbType.Decimal);
			parameters.Add("@funb_total_credits", 0m, DbType.Decimal);
			parameters.Add("@funb_total_debits", 0m, DbType.Decimal);
			parameters.Add("@ledger_bal", 0m, DbType.Decimal);
			parameters.Add("@update_status", "I");
			parameters.Add("@bai_file_id_OUTPUT", dbType: DbType.String, size: 14, direction: ParameterDirection.Output);
			parameters.Add("@proc_result", dbType: DbType.Int32, direction: ParameterDirection.ReturnValue);

			await connection.ExecuteAsync(
				"up_iud_BAI_File_Summary",
				parameters,
				commandType: CommandType.StoredProcedure
			);

			var procResult = parameters.Get<int>("@proc_result");
			var baiFileId = parameters.Get<string>("@bai_file_id_OUTPUT");
			return (procResult, baiFileId);
		}

		public async Task UpdateBaiFileSummaryAsync(long fileIdNum, string baiFileId, decimal availableBal, decimal collectedBal, decimal totalCredits, decimal totalDebits, decimal ledgerBal, DateTime baiFileDate)
		{
			using var connection = _dbContext.CreateConnection("dds_schema");
			var parameters = new DynamicParameters();
			parameters.Add("@file_id_num", fileIdNum, DbType.Int64);
			parameters.Add("@bai_file_id", decimal.Parse(baiFileId));
			parameters.Add("@available_bal", availableBal);
			parameters.Add("@collected_bal", collectedBal);
			parameters.Add("@funb_total_credits", totalCredits);
			parameters.Add("@funb_total_debits", totalDebits);
			parameters.Add("@ledger_bal", ledgerBal);
			parameters.Add("@bai_file_datetime", baiFileDate.ToString("yyyy-MM-dd"));
			parameters.Add("@update_status", "U");
			parameters.Add("@created_by", null);

			await connection.ExecuteAsync("up_iud_BAI_File_Summary", parameters, commandType: CommandType.StoredProcedure);
		}

		public async Task InsertWorkFileAsync(WorkFileInsert data)
		{
			using var connection = _dbContext.CreateConnection("dds_schema");
			var p = new DynamicParameters();
			p.Add("@record_id", null, DbType.Decimal);
			p.Add("@invalid_record_id", null, DbType.Decimal);
			p.Add("@bai_file_id", data.BaiFileId);
			p.Add("@tot_funb_benefit_amt", data.TotalFunbBenefitAmount);
			p.Add("@dr_cr_flag", data.DrCrFlag);
			p.Add("@as_of_datetime", data.AsOfDate.ToString("yyyy-MM-dd"));
			p.Add("@created_by", data.CreatedBy);
			p.Add("@record_status", "A");
			p.Add("@shared_dd_num_ind", "N");
			p.Add("@dd_num", data.DdNumber ?? "");
			p.Add("@institution_code", data.InstitutionCode);
			p.Add("@affinity_acct_num", data.AffinityAccountNumber);
			p.Add("@medical_record_num", data.MedicalRecordNumber);
			p.Add("@income_source_type", data.IncomeSourceType);
			p.Add("@name", data.Name);
			p.Add("@deceased_ind", "N");
			p.Add("@discharge_date", data.DischargeDate);
			p.Add("@comment", data.Comment);
			p.Add("@pa_posting_status", data.PaPostingStatus);
			p.Add("@pf_posting_status", data.PfPostingStatus);
			p.Add("@pa_err_code", data.PaErrCode, DbType.Decimal);
			p.Add("@pf_err_code", data.PfErrCode, DbType.Decimal);
			p.Add("@validated", "N");
			p.Add("@update_status", "I");

			await connection.ExecuteAsync("up_iud_DDWorkFile", p, commandType: CommandType.StoredProcedure);
		}

		public async Task<bool> AnyWorkFileRecordsAsync()
		{
			const string sql = "SELECT 1 FROM DD_WORK_FILE WHERE VALIDATED = 'N'";
			try
			{
				using var connection = _dbContext.CreateConnection("dds_schema");
				var result = await connection.QueryFirstOrDefaultAsync<int?>(sql);
				return result.HasValue;
			}
			catch (Exception ex)
			{
				_logger.LogError(ex, "Error checking for records in DD_WORK_FILE");
				throw;
			}
		}
	}

	public class WorkFileInsert
	{
		public decimal BaiFileId { get; set; }
		public decimal TotalFunbBenefitAmount { get; set; }
		public string DrCrFlag { get; set; } = string.Empty;
		public DateTime AsOfDate { get; set; }
		public string CreatedBy { get; set; } = string.Empty;
		public string? DdNumber { get; set; }
		public string? InstitutionCode { get; set; }
		public string? AffinityAccountNumber { get; set; }
		public string? MedicalRecordNumber { get; set; }
		public string IncomeSourceType { get; set; } = string.Empty;
		public string? Name { get; set; }
		public DateTime? DischargeDate { get; set; }
		public string Comment { get; set; } = string.Empty;
		public string? PaPostingStatus { get; set; }
		public string? PfPostingStatus { get; set; }
		public decimal? PaErrCode { get; set; }
		public decimal? PfErrCode { get; set; }
	}
}


