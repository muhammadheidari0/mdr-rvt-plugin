using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Autodesk.Revit.UI;
using Mdr.Revit.Client.Http;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;
using Mdr.Revit.Core.UseCases;
using Mdr.Revit.Infra.Logging;
using Mdr.Revit.RevitAdapter.Extractors;

namespace Mdr.Revit.Addin.Commands
{
    public sealed class PublishSheetsCommand
    {
        private readonly Func<Uri, IApiClient> _apiClientFactory;
        private readonly IRevitExtractor _revitExtractor;
        private readonly PdfExporter _pdfExporter;
        private readonly NativeExporter _nativeExporter;
        private readonly PluginLogger _logger;

        public PublishSheetsCommand()
            : this(
                baseAddress => new ApiClient(baseAddress),
                new RevitExtractorAdapter(new SheetExtractor(), new ScheduleExtractor()),
                new PdfExporter(),
                new NativeExporter(),
                new PluginLogger(DefaultLogDirectory()))
        {
        }

        public PublishSheetsCommand(UIDocument uiDocument)
            : this(
                baseAddress => new ApiClient(baseAddress),
                RevitApiExtractors.CreateExtractor(uiDocument),
                RevitApiExtractors.CreatePdfExporter(uiDocument),
                RevitApiExtractors.CreateNativeExporter(uiDocument),
                new PluginLogger(DefaultLogDirectory()))
        {
        }

        internal PublishSheetsCommand(
            Func<Uri, IApiClient> apiClientFactory,
            IRevitExtractor revitExtractor,
            PdfExporter pdfExporter,
            NativeExporter nativeExporter,
            PluginLogger logger)
        {
            _apiClientFactory = apiClientFactory ?? throw new ArgumentNullException(nameof(apiClientFactory));
            _revitExtractor = revitExtractor ?? throw new ArgumentNullException(nameof(revitExtractor));
            _pdfExporter = pdfExporter ?? throw new ArgumentNullException(nameof(pdfExporter));
            _nativeExporter = nativeExporter ?? throw new ArgumentNullException(nameof(nativeExporter));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        public string Id => "mdr.publishSheets";

        public IReadOnlyList<PublishSheetItem> GetSelectedSheets()
        {
            IReadOnlyList<PublishSheetItem> extracted = _revitExtractor.GetSelectedSheets();
            if (extracted == null || extracted.Count == 0)
            {
                return Array.Empty<PublishSheetItem>();
            }

            List<PublishSheetItem> normalized = new List<PublishSheetItem>(extracted.Count);
            for (int i = 0; i < extracted.Count; i++)
            {
                PublishSheetItem source = extracted[i];
                if (source == null)
                {
                    continue;
                }

                PublishSheetItem item = source.Clone();
                item.ItemIndex = normalized.Count;
                normalized.Add(item);
            }

            return normalized;
        }

        public async Task<PublishSheetsCommandResult> ExecuteAsync(
            PublishSheetsCommandRequest request,
            CancellationToken cancellationToken)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            ValidateRequest(request);

            Uri baseAddress = new Uri(request.BaseUrl, UriKind.Absolute);
            IApiClient apiClient = _apiClientFactory(baseAddress);

            PublishSheetsUseCase useCase = new PublishSheetsUseCase(apiClient, _revitExtractor);
            PublishBatchRequest publishRequest = BuildPublishRequest(request);
            CorrelationContext.CurrentRunUid = publishRequest.RunClientId;

            try
            {
                _logger.Info("Starting sheet publish run.");

                await apiClient
                    .LoginAsync(request.Username, request.Password, cancellationToken)
                    .ConfigureAwait(false);
                _logger.Info("Login succeeded.");

                PublishBatchResponse initialResponse = await useCase
                    .ExecuteAsync(publishRequest, cancellationToken)
                    .ConfigureAwait(false);

                PublishBatchResponse? retryResponse = null;
                if (request.RetryFailedItems && initialResponse.Summary.FailedCount > 0)
                {
                    _logger.Info("Retrying failed publish items.");
                    retryResponse = await useCase
                        .RetryFailedAsync(publishRequest, initialResponse, cancellationToken)
                        .ConfigureAwait(false);
                }

                PublishBatchResponse finalResponse = retryResponse ?? initialResponse;
                _logger.Info(
                    "Publish finished with status=" + finalResponse.Summary.Status +
                    " success=" + finalResponse.Summary.SuccessCount +
                    " failed=" + finalResponse.Summary.FailedCount +
                    " duplicate=" + finalResponse.Summary.DuplicateCount);

                return new PublishSheetsCommandResult
                {
                    InitialResponse = initialResponse,
                    RetryResponse = retryResponse,
                    OutputDirectory = ResolveOutputDirectory(request),
                };
            }
            finally
            {
                CorrelationContext.CurrentRunUid = string.Empty;

                if (apiClient is IDisposable disposable)
                {
                    disposable.Dispose();
                }
            }
        }

        private PublishBatchRequest BuildPublishRequest(PublishSheetsCommandRequest request)
        {
            PublishBatchRequest publishRequest = new PublishBatchRequest
            {
                ProjectCode = request.ProjectCode,
                ModelGuid = request.ModelGuid ?? string.Empty,
                ModelTitle = request.ModelTitle ?? string.Empty,
                RevitVersion = string.IsNullOrWhiteSpace(request.RevitVersion) ? "2026" : request.RevitVersion,
                PluginVersion = string.IsNullOrWhiteSpace(request.PluginVersion) ? "0.3.0" : request.PluginVersion,
            };

            IReadOnlyList<PublishSheetItem> inputItems = request.Items.Count > 0
                ? request.Items
                : GetSelectedSheets();

            foreach (PublishSheetItem item in inputItems)
            {
                if (item == null)
                {
                    continue;
                }

                PublishSheetItem clone = item.Clone();
                if (string.IsNullOrWhiteSpace(clone.StatusCode))
                {
                    clone.StatusCode = request.DefaultStatusCode;
                }
                clone.IncludeNative = request.IncludeNative || clone.IncludeNative;
                publishRequest.Items.Add(clone);
            }

            AttachExportedArtifacts(publishRequest.Items, ResolveOutputDirectory(request), request.IncludeNative);
            return publishRequest;
        }

        private void AttachExportedArtifacts(
            IReadOnlyList<PublishSheetItem> items,
            string outputDirectory,
            bool includeNative)
        {
            if (items.Count == 0)
            {
                return;
            }

            bool needsPdf = false;
            bool needsNative = false;
            for (int i = 0; i < items.Count; i++)
            {
                PublishSheetItem item = items[i];
                if (string.IsNullOrWhiteSpace(item.PdfFilePath))
                {
                    needsPdf = true;
                }

                if (includeNative && item.IncludeNative && string.IsNullOrWhiteSpace(item.NativeFilePath))
                {
                    needsNative = true;
                }
            }

            if (!needsPdf && !needsNative)
            {
                return;
            }

            List<string> sheetIds = new List<string>(items.Count);
            for (int i = 0; i < items.Count; i++)
            {
                string sheetId = string.IsNullOrWhiteSpace(items[i].SheetUniqueId)
                    ? ("sheet_" + i)
                    : items[i].SheetUniqueId;
                sheetIds.Add(sheetId);
            }

            IReadOnlyList<string> pdfFiles = needsPdf
                ? _pdfExporter.ExportSheetsToPdf(sheetIds, outputDirectory)
                : Array.Empty<string>();
            IReadOnlyList<string> nativeFiles = needsNative
                ? _nativeExporter.ExportNativeFiles(sheetIds, outputDirectory)
                : Array.Empty<string>();

            for (int i = 0; i < items.Count; i++)
            {
                PublishSheetItem item = items[i];
                if (string.IsNullOrWhiteSpace(item.PdfFilePath) && i < pdfFiles.Count)
                {
                    item.PdfFilePath = pdfFiles[i];
                }

                if (item.IncludeNative && string.IsNullOrWhiteSpace(item.NativeFilePath) && i < nativeFiles.Count)
                {
                    item.NativeFilePath = nativeFiles[i];
                }
            }
        }

        private static string ResolveOutputDirectory(PublishSheetsCommandRequest request)
        {
            if (!string.IsNullOrWhiteSpace(request.OutputDirectory))
            {
                Directory.CreateDirectory(request.OutputDirectory);
                return request.OutputDirectory;
            }

            string path = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "MDR",
                "RevitPlugin",
                "publish");
            Directory.CreateDirectory(path);
            return path;
        }

        private static void ValidateRequest(PublishSheetsCommandRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.BaseUrl))
            {
                throw new InvalidOperationException("BaseUrl is required.");
            }

            if (string.IsNullOrWhiteSpace(request.Username))
            {
                throw new InvalidOperationException("Username is required.");
            }

            if (string.IsNullOrWhiteSpace(request.Password))
            {
                throw new InvalidOperationException("Password is required.");
            }

            if (string.IsNullOrWhiteSpace(request.ProjectCode))
            {
                throw new InvalidOperationException("ProjectCode is required.");
            }
        }

        private static string DefaultLogDirectory()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "MDR",
                "RevitPlugin",
                "logs");
        }
    }

    public sealed class PublishSheetsCommandRequest
    {
        public string BaseUrl { get; set; } = "http://127.0.0.1:8000";

        public string Username { get; set; } = string.Empty;

        public string Password { get; set; } = string.Empty;

        public string ProjectCode { get; set; } = string.Empty;

        public string ModelGuid { get; set; } = string.Empty;

        public string ModelTitle { get; set; } = string.Empty;

        public string RevitVersion { get; set; } = "2026";

        public string PluginVersion { get; set; } = "0.3.0";

        public string DefaultStatusCode { get; set; } = "IFA";

        public bool IncludeNative { get; set; }

        public bool RetryFailedItems { get; set; } = true;

        public string OutputDirectory { get; set; } = string.Empty;

        public List<PublishSheetItem> Items { get; } = new List<PublishSheetItem>();
    }

    public sealed class PublishSheetsCommandResult
    {
        public PublishBatchResponse InitialResponse { get; set; } = new PublishBatchResponse();

        public PublishBatchResponse? RetryResponse { get; set; }

        public string OutputDirectory { get; set; } = string.Empty;

        public PublishBatchResponse FinalResponse
        {
            get
            {
                return RetryResponse ?? InitialResponse;
            }
        }
    }
}
