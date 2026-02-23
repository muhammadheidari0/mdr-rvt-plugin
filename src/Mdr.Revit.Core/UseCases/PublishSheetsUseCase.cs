using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;
using Mdr.Revit.Core.Validation;

namespace Mdr.Revit.Core.UseCases
{
    public sealed class PublishSheetsUseCase
    {
        private readonly IApiClient _apiClient;
        private readonly IRevitExtractor _revitExtractor;

        public PublishSheetsUseCase(IApiClient apiClient, IRevitExtractor revitExtractor)
        {
            _apiClient = apiClient ?? throw new ArgumentNullException(nameof(apiClient));
            _revitExtractor = revitExtractor ?? throw new ArgumentNullException(nameof(revitExtractor));
        }

        public async Task<PublishBatchResponse> ExecuteAsync(
            PublishBatchRequest request,
            CancellationToken cancellationToken)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            if (request.Items.Count == 0)
            {
                IReadOnlyList<PublishSheetItem> selectedSheets = _revitExtractor.GetSelectedSheets();
                foreach (PublishSheetItem item in selectedSheets)
                {
                    request.Items.Add(item);
                }
            }

            NormalizeItems(request.Items);
            HydrateFileHashes(request.Items);
            request.BuildFilesManifest();
            BusinessRules.EnsurePublishRequestIsValid(request);

            PublishBatchResponse response = await _apiClient
                .PublishBatchAsync(request, cancellationToken)
                .ConfigureAwait(false);

            return response;
        }

        public async Task<PublishBatchResponse> RetryFailedAsync(
            PublishBatchRequest originalRequest,
            PublishBatchResponse previousResponse,
            CancellationToken cancellationToken)
        {
            if (originalRequest == null)
            {
                throw new ArgumentNullException(nameof(originalRequest));
            }

            if (previousResponse == null)
            {
                throw new ArgumentNullException(nameof(previousResponse));
            }

            PublishBatchRequest retryRequest = originalRequest.CreateRetryRequest(previousResponse);
            if (retryRequest.Items.Count == 0)
            {
                return new PublishBatchResponse
                {
                    RunId = previousResponse.RunId,
                    Summary = new PublishBatchSummary
                    {
                        RequestedCount = 0,
                        SuccessCount = 0,
                        FailedCount = 0,
                        DuplicateCount = 0,
                        Status = "no_retry_needed",
                    },
                };
            }

            return await ExecuteAsync(retryRequest, cancellationToken).ConfigureAwait(false);
        }

        private static void NormalizeItems(List<PublishSheetItem> items)
        {
            HashSet<int> usedIndexes = new HashSet<int>();
            for (int i = 0; i < items.Count; i++)
            {
                PublishSheetItem item = items[i];
                if (item == null)
                {
                    continue;
                }

                int normalized = item.ItemIndex;
                if (normalized < 0 || usedIndexes.Contains(normalized))
                {
                    normalized = i;
                }

                while (usedIndexes.Contains(normalized))
                {
                    normalized++;
                }

                item.ItemIndex = normalized;
                usedIndexes.Add(normalized);
            }
        }

        private static void HydrateFileHashes(IReadOnlyList<PublishSheetItem> items)
        {
            if (items == null)
            {
                return;
            }

            for (int i = 0; i < items.Count; i++)
            {
                PublishSheetItem item = items[i];
                if (item == null)
                {
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(item.FileSha256))
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(item.PdfFilePath))
                {
                    continue;
                }

                if (!File.Exists(item.PdfFilePath))
                {
                    continue;
                }

                item.FileSha256 = ComputeSha256(item.PdfFilePath);
            }
        }

        private static string ComputeSha256(string filePath)
        {
            using (SHA256 sha = SHA256.Create())
            using (FileStream stream = File.OpenRead(filePath))
            {
                byte[] hash = sha.ComputeHash(stream);
                StringBuilder builder = new StringBuilder(hash.Length * 2);
                for (int i = 0; i < hash.Length; i++)
                {
                    builder.Append(hash[i].ToString("x2"));
                }

                return builder.ToString();
            }
        }
    }
}
