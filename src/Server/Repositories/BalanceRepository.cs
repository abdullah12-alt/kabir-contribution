using Dapper;
using DDS.API.Data;
using Server.Models;
using System.Data;
using System.Threading.Tasks;
using System.Collections.Generic;
namespace Server.Repositories
{
    public interface IBalanceRepository
    {
        Task<BalanceSummaryDto> GetBalanceSummaryAsync();
        Task<IList<InvalidRecord>> GetInvalidRecordsAsync();
        Task<int> InsertBalanceAsync(BalanceInsertDto dto);
        Task<IList<SummaryRecordDto>> GetSummaryRecordsAsync();
    }
    public class BalanceRepository : IBalanceRepository
    {
        private readonly DapperDbContext _dbContext;
        private readonly ILogger<BalanceRepository> _logger;

        public BalanceRepository(DapperDbContext dbContext, ILogger<BalanceRepository> logger)
        {
            _dbContext = dbContext;
            _logger = logger;
        }
        public async Task<BalanceSummaryDto> GetBalanceSummaryAsync()
        {
            using var connection = _dbContext.CreateConnection("dds_schema");

            // 1. Get latest balance
            var latestBalance = await connection.QueryFirstOrDefaultAsync<BalanceRecord>(
                "SELECT TOP 1 * FROM DD_BALANCE ORDER BY BALANCE_ID DESC");

            // 2. Determine last balance date
            var dtLastBalance = latestBalance?.CREATED_DATETIME ?? DateTime.Now.AddYears(-1);

            // 3. Get carryover balance
            decimal carryOverBal = 0;
            if (latestBalance != null)
                carryOverBal = (latestBalance.TOT_CR_DR_PRE_EDIT_RPT ?? 0) + (latestBalance.TOT_CR_DR_DECEASED_EXCEPT ?? 0);

            // 4. Sum hidden records since last balance
            var dTotHidden = await GetHiddenAdjustmentsAsync(connection, dtLastBalance);

            // 5. Sum posted debits/credits since last balance
            var dTotBenefit = await GetPostedBenefitsAsync(connection, dtLastBalance);

            // 6. Get ledger balance
            var ledgerBal = await connection.QueryFirstOrDefaultAsync<decimal?>(
                "SELECT TOP 1 LEDGER_BAL FROM DD_BAI_FILE_SUMMARY ORDER BY FILE_ID_NUM DESC") ?? 0;

            // 7. Sum current invalid and deceased records
            var invalids = await connection.QueryAsync<InvalidRecord>(
                "SELECT * FROM DD_INVALID_REC WHERE RECORD_STATUS = 'A'");
            decimal mdInvalidTotal = 0, mdDeceasedTotal = 0;
            foreach (var rec in invalids)
            {
                if (rec.DECEASED_IND == "N")
                    mdInvalidTotal += rec.DR_CR_FLAG == "DR" ? -rec.TOT_FUNB_BENEFIT_AMT : rec.TOT_FUNB_BENEFIT_AMT;
                else
                    mdDeceasedTotal += rec.DR_CR_FLAG == "DR" ? -rec.TOT_FUNB_BENEFIT_AMT : rec.TOT_FUNB_BENEFIT_AMT;
            }
            // 8. Determine beginning balance
            // Note: User can supply this, otherwise use last ADJ_ENDING_BAL
            decimal beginningBalance = latestBalance?.ADJ_ENDING_BAL ?? 0;

            // 9. Compute totals
            var endingBalance = beginningBalance - carryOverBal + dTotBenefit + dTotHidden;
            var adjustedEndingBalance = endingBalance + mdInvalidTotal + mdDeceasedTotal;
            var difference = Math.Round(adjustedEndingBalance - ledgerBal, 2);
            // 10. Return all in a summary DTO
            return new BalanceSummaryDto
            {
                BeginningBalance = beginningBalance,
                CarryOverBal = carryOverBal,
                LastBalanceDate = dtLastBalance,
                HiddenAdjustments = dTotHidden,
                PostedBenefits = dTotBenefit,
                LedgerBalance = ledgerBal,
                InvalidTotal = mdInvalidTotal,
                DeceasedTotal = mdDeceasedTotal,
                EndingBalance = endingBalance,
                AdjustedEndingBalance = adjustedEndingBalance,
                Difference = difference
            };
        }
        private async Task<decimal> GetHiddenAdjustmentsAsync(IDbConnection connection, DateTime dtLastBalance)
        {
            decimal dTotHidden = 0;
            var dr = await connection.QueryFirstOrDefaultAsync<decimal?>(
                "SELECT SUM(TOT_FUNB_BENEFIT_AMT) FROM DD_INVALID_REC WHERE DR_CR_FLAG = 'DR' AND LAST_MOD_DATETIME >= @dt AND RECORD_STATUS = 'I'",
                new { dt = dtLastBalance });
            if (dr.HasValue) dTotHidden = -dr.Value;

            var cr = await connection.QueryFirstOrDefaultAsync<decimal?>(
                "SELECT SUM(TOT_FUNB_BENEFIT_AMT) FROM DD_INVALID_REC WHERE DR_CR_FLAG = 'CR' AND LAST_MOD_DATETIME >= @dt AND RECORD_STATUS = 'I'",
                new { dt = dtLastBalance });
            if (cr.HasValue) dTotHidden += cr.Value;

            return dTotHidden;
        }

        private async Task<decimal> GetPostedBenefitsAsync(IDbConnection connection, DateTime dtLastBalance)
        {
            decimal dTotBenefit = 0;
            var dr = await connection.QueryFirstOrDefaultAsync<decimal?>(
                "SELECT SUM(TOT_FUNB_BENEFIT_AMT) FROM DD_POSTING_HISTORY WHERE DR_CR_FLAG = 'DR' AND POSTED_DATETIME > @dt",
                new { dt = dtLastBalance });
            if (dr.HasValue) dTotBenefit = -dr.Value;

            var cr = await connection.QueryFirstOrDefaultAsync<decimal?>(
                "SELECT SUM(TOT_FUNB_BENEFIT_AMT) FROM DD_POSTING_HISTORY WHERE DR_CR_FLAG = 'CR' AND POSTED_DATETIME > @dt",
                new { dt = dtLastBalance });
            if (cr.HasValue) dTotBenefit += cr.Value;

            return dTotBenefit;
        }


        public async Task<IList<InvalidRecord>> GetInvalidRecordsAsync()
        {
            using var connection = _dbContext.CreateConnection("dds_schema");
            var result = await connection.QueryAsync<InvalidRecord>(
                "SELECT * FROM DD_INVALID_REC WHERE RECORD_STATUS = 'A'");
            return result.AsList();
        }
        public async Task<int> InsertBalanceAsync(BalanceInsertDto dto)
        {
            const string proc = "up_i_Balance";
            var parameters = new DynamicParameters();
            parameters.Add("beginning_bal", dto.BeginningBalance);
            parameters.Add("carryover_bal", dto.CarryOverBalance);
            parameters.Add("tot_cr_dr_posted", dto.TotalPosted);
            parameters.Add("ending_bal", dto.EndingBalance);
            parameters.Add("tot_cr_dr_pre_edit_rpt", dto.InvalidTotal);
            parameters.Add("tot_cr_dr_deceased_except", dto.DeceasedTotal);
            parameters.Add("adj_ending_bal", dto.AdjustedEndingBalance);
            parameters.Add("ledger_bal", dto.LedgerBalance);
            parameters.Add("created_by", dto.CreatedBy);
            parameters.Add("tot_cr_dr_adjustments", dto.Adjustments);
            parameters.Add("RETURN_VALUE", dbType: DbType.Int32, direction: ParameterDirection.ReturnValue);

            using var connection = _dbContext.CreateConnection("dds_schema");
            await connection.ExecuteAsync(proc, parameters, commandType: CommandType.StoredProcedure);
            return parameters.Get<int>("RETURN_VALUE");
        }

        public async Task<IList<SummaryRecordDto>> GetSummaryRecordsAsync()
        {
            using var connection = _dbContext.CreateConnection("dds_schema");
            var sql = @"SELECT BAI_FILE_DATETIME, LEDGER_BAL, AVAILABLE_BAL, COLLECTED_BAL, 
                    FUNB_TOTAL_CREDITS, FUNB_TOTAL_DEBITS, CREATED_DATETIME, CREATED_BY
                    FROM DD_BAI_FILE_SUMMARY
                    WHERE BAI_FILE_DATETIME >= @date
                    ORDER BY FILE_ID_NUM DESC";
            var date = DateTime.Now.AddDays(-10).Date;
            var result = await connection.QueryAsync<SummaryRecordDto>(sql, new { date });
            return result.AsList();
        }
    }
}

