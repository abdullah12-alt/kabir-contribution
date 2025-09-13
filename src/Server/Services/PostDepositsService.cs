// PostDepositsService.cs
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Server.Models;
using Server.Repositories;

namespace Server.Services
{
    public interface IPostDepositsService
    {
        Task<PostResult> PostDepositsAsync(string userId, string flatFileDirectory, string processingId);
        Task<PostCounts> GetPostCountsAsync();
        Task<IEnumerable<PostingAck>> GetAllPostingAcksAsync();
    }

    public class PostDepositsService : IPostDepositsService
    {
        private readonly IPostDepositsRepository _repo;
        private readonly ILogger<PostDepositsService> _logger;

        public PostDepositsService(IPostDepositsRepository repo, ILogger<PostDepositsService> logger)
        {
            _repo = repo;
            _logger = logger;
        }

        public async Task<PostResult> PostDepositsAsync(string userId, string flatFileDirectory, string processingId)
        {
           

            // 2. Ensure ToPost directory exists and is empty
            await _repo.EnsureDirectoryExistsAsync(flatFileDirectory);
            await _repo.DeleteAllFilesInDirectoryAsync(flatFileDirectory);

            // 3. Get counts
            var counts = await _repo.GetPostCountsAsync();
            if (counts.TotalRecords == 0)
                return new PostResult { Success = false, Message = "No records to post." };

            // 4. Get config info
            var config = await _repo.GetConfigInfoAsync();

            // 5. Get records to post to Affinity
            var records = await _repo.GetRecordsToPostToAffinityAsync();

            // 6. Create HL7 files for each record
            var createdFiles = new List<string>();
            var transactionDate = DateTime.Now;

            // Get a transaction group number for this batch
            var transGroupNum = await _repo.GetTransactionGroupNumberAsync();

            foreach (var rec in records)
            {
                var file = await _repo.CreateHL7FlatFileAsync(rec, flatFileDirectory, config, processingId, transactionDate);
                createdFiles.Add(file);
                bool marked = await _repo.MarkAsSentForPostingAsync(rec.VALID_RECORD_ID);
                if (!marked)
                {
                    // Handle error (log, retry, etc.)
                }


                if (rec.PA_DISTRIBUTION_AMT > 0)
                {
                    var (pfSuccess, pfStatus) = await _repo.PostPFTransactionAsync(
                        rec.VALID_RECORD_ID,
                        transGroupNum,
                        userId,
                        "Success");

                }
            }
                return new PostResult
            {
                Success = true,
                Message = "Files created and ready for Affinity transfer.",
                TotalRecords = counts.TotalRecords,
                TotalPATrans = counts.TotalPATrans,
                TotalPFTrans = counts.TotalPFTrans,
                CreatedFiles = createdFiles
            };
        }
        public async Task<PostCounts> GetPostCountsAsync()
        {
            var counts = await _repo.GetPostCountsAsync();
            return counts;
        }
        public async Task<IEnumerable<PostingAck>> GetAllPostingAcksAsync()
        {
            return await _repo.GetAllPostingAcksAsync();
        }
    }

  
public class PostResult
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public int TotalRecords { get; set; }
        public int TotalPATrans { get; set; }
        public int TotalPFTrans { get; set; }
        public List<string> CreatedFiles { get; set; }
    }
}


