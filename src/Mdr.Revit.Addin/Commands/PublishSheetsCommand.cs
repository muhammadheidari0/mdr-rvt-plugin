using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Autodesk.Revit.UI;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;
using Mdr.Revit.Core.UseCases;
using Mdr.Revit.Infra.Logging;
using Mdr.Revit.RevitAdapter.Extractors;

namespace Mdr.Revit.Addin.Commands
{
    public sealed class PublishSheetsCommand
    {
        private readonly Func<ApiClientFactoryOptions, IApiClient> _apiClientFactory;
        private readonly IRevitExtractor _revitExtractor;
        private readonly PdfExporter _pdfExporter;
        private readonly NativeExporter _nativeExporter;
        private readonly PluginLogger _logger;

        public PublishSheetsCommand()
            : this(
                ApiClientFactory.Create,
                new RevitExtractorAdapter(new SheetExtractor(), new ScheduleExtractor()),
                new PdfExporter(),
                new NativeExporter(),
                new PluginLogger(DefaultLogDirectory()))
        {
        }

        public PublishSheetsCommand(UIDocument uiDocument)
            : this(
                ApiClientFactory.Create,
                RevitApiExtractors.CreateExtractor(uiDocument),
                RevitApiExtractors.CreatePdfExporter(uiDocument),
                RevitApiExtractors.CreateNativeExporter(uiDocument),
                new PluginLogger(DefaultLogDirectory()))
        {
        }

        internal PublishSheetsCommand(
            Func<ApiClientFactoryOptions, IApiClient> apiClientFactory,
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

            IApiClient apiClient = _apiClientFactory(new ApiClientFactoryOptions
            {
                BaseAddress = new Uri(request.BaseUrl, UriKind.Absolute),
                RequestTimeoutSeconds = request.RequestTimeoutSeconds,
                AllowInsecureTls = request.AllowInsecureTls,
            });
            PublishSheetsUseCase useCase = new PublishSheetsUseCase(apiClient, _revitExtractor);

            try
            {
                _logger.Info("Starting sheet publish run.");

                await ExecuteLoginPreflightAsync(apiClient, request, cancellationToken).ConfigureAwait(false);
                _logger.Info("Login succeeded.");

                PublishBuildResult buildResult = BuildPublishRequest(request);
                CorrelationContext.CurrentRunUid = buildResult.PublishRequest.RunClientId;

                PublishBatchResponse initialServerResponse = buildResult.PublishRequest.Items.Count > 0
                    ? await useCase.ExecuteAsync(buildResult.PublishRequest, cancellationToken).ConfigureAwait(false)
                    : new PublishBatchResponse
                    {
                        RunId = buildResult.PublishRequest.RunClientId,
                        Summary = new PublishBatchSummary
                        {
                            RequestedCount = 0,
                            SuccessCount = 0,
                            FailedCount = 0,
                            DuplicateCount = 0,
                            Status = "no_publishable_items",
                        },
                    };

                PublishBatchResponse? retryServerResponse = null;
                if (request.RetryFailedItems &&
                    buildResult.PublishRequest.Items.Count > 0 &&
                    initialServerResponse.Summary.FailedCount > 0)
                {
                    _logger.Info("Retrying failed publish items.");
                    retryServerResponse = await useCase
                        .RetryFailedAsync(buildResult.PublishRequest, initialServerResponse, cancellationToken)
                        .ConfigureAwait(false);
                }

                PublishBatchResponse initialResponse = MergeWithLocalFailures(buildResult, initialServerResponse);
                PublishBatchResponse? retryResponse = retryServerResponse == null
                    ? null
                    : MergeWithLocalFailures(buildResult, retryServerResponse);
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
                    OutputDirectory = buildResult.OutputDirectory,
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

        private PublishBuildResult BuildPublishRequest(PublishSheetsCommandRequest request)
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

            List<int> requestedItemIndexes = new List<int>(inputItems.Count);
            for (int i = 0; i < inputItems.Count; i++)
            {
                PublishSheetItem item = inputItems[i];
                if (item == null)
                {
                    continue;
                }

                PublishSheetItem clone = item.Clone();
                if (clone.ItemIndex < 0)
                {
                    clone.ItemIndex = i;
                }

                if (string.IsNullOrWhiteSpace(clone.StatusCode))
                {
                    clone.StatusCode = request.DefaultStatusCode;
                }

                clone.IncludeNative = request.IncludeNative || clone.IncludeNative;
                publishRequest.Items.Add(clone);
                requestedItemIndexes.Add(clone.ItemIndex);
            }

            string outputDirectory = ResolveOutputDirectory(request);
            Dictionary<int, PublishItemResult> localFailures = AttachExportedArtifacts(
                publishRequest.Items,
                outputDirectory,
                request.IncludeNative,
                request.NativeFormat);

            for (int i = publishRequest.Items.Count - 1; i >= 0; i--)
            {
                PublishSheetItem item = publishRequest.Items[i];
                if (localFailures.ContainsKey(item.ItemIndex))
                {
                    publishRequest.Items.RemoveAt(i);
                }
            }

            return new PublishBuildResult
            {
                PublishRequest = publishRequest,
                RequestedItemIndexes = requestedItemIndexes,
                LocalFailures = localFailures,
                OutputDirectory = outputDirectory,
            };
        }

        private Dictionary<int, PublishItemResult> AttachExportedArtifacts(
            IReadOnlyList<PublishSheetItem> items,
            string outputDirectory,
            bool includeNative,
            string nativeFormat)
        {
            Dictionary<int, PublishItemResult> localFailures = new Dictionary<int, PublishItemResult>();
            if (items.Count == 0)
            {
                return localFailures;
            }

            List<PublishSheetItem> pdfTargets = new List<PublishSheetItem>();
            List<PublishSheetItem> nativeTargets = new List<PublishSheetItem>();
            for (int i = 0; i < items.Count; i++)
            {
                PublishSheetItem item = items[i];
                if (string.IsNullOrWhiteSpace(item.PdfFilePath))
                {
                    pdfTargets.Add(item);
                }

                if (includeNative && item.IncludeNative && string.IsNullOrWhiteSpace(item.NativeFilePath))
                {
                    nativeTargets.Add(item);
                }
            }

            Dictionary<int, ExportArtifact> pdfByItem = BuildArtifactMap(
                pdfTargets.Count == 0
                    ? Array.Empty<ExportArtifact>()
                    : _pdfExporter.ExportSheetsToPdf(pdfTargets, outputDirectory),
                ExportArtifactKinds.Pdf);

            string normalizedNativeFormat = string.IsNullOrWhiteSpace(nativeFormat)
                ? "dwg"
                : nativeFormat.Trim().ToLowerInvariant();
            Dictionary<int, ExportArtifact> nativeByItem;
            if (nativeTargets.Count > 0 && !string.Equals(normalizedNativeFormat, "dwg", StringComparison.OrdinalIgnoreCase))
            {
                nativeByItem = new Dictionary<int, ExportArtifact>();
                for (int i = 0; i < nativeTargets.Count; i++)
                {
                    PublishSheetItem item = nativeTargets[i];
                    nativeByItem[item.ItemIndex] = new ExportArtifact
                    {
                        ItemIndex = item.ItemIndex,
                        SheetUniqueId = item.SheetUniqueId,
                        Kind = ExportArtifactKinds.Native,
                        ErrorCode = "native_format_unsupported",
                        ErrorMessage = "Native format '" + normalizedNativeFormat + "' is not supported. Use 'dwg'.",
                    };
                }
            }
            else
            {
                nativeByItem = BuildArtifactMap(
                    nativeTargets.Count == 0
                        ? Array.Empty<ExportArtifact>()
                        : _nativeExporter.ExportNativeFiles(nativeTargets, outputDirectory),
                    ExportArtifactKinds.Native);
            }

            for (int i = 0; i < items.Count; i++)
            {
                PublishSheetItem item = items[i];
                int itemIndex = item.ItemIndex;

                if (string.IsNullOrWhiteSpace(item.PdfFilePath))
                {
                    if (!pdfByItem.TryGetValue(itemIndex, out ExportArtifact? artifact))
                    {
                        RegisterLocalFailure(localFailures, item, "export_pdf_failed", "PDF export output was not produced.");
                        continue;
                    }

                    if (!artifact.IsSuccess())
                    {
                        RegisterLocalFailure(localFailures, item, NormalizeErrorCode(artifact.ErrorCode, "export_pdf_failed"), artifact.ErrorMessage);
                        continue;
                    }

                    item.PdfFilePath = artifact.FilePath;
                    if (string.IsNullOrWhiteSpace(item.FileSha256) && !string.IsNullOrWhiteSpace(artifact.FileSha256))
                    {
                        item.FileSha256 = artifact.FileSha256;
                    }
                }

                if (includeNative && item.IncludeNative && string.IsNullOrWhiteSpace(item.NativeFilePath))
                {
                    if (!nativeByItem.TryGetValue(itemIndex, out ExportArtifact? artifact))
                    {
                        RegisterLocalFailure(localFailures, item, "export_native_failed", "Native export output was not produced.");
                        continue;
                    }

                    if (!artifact.IsSuccess())
                    {
                        RegisterLocalFailure(localFailures, item, NormalizeErrorCode(artifact.ErrorCode, "export_native_failed"), artifact.ErrorMessage);
                        continue;
                    }

                    item.NativeFilePath = artifact.FilePath;
                }

                if (localFailures.ContainsKey(itemIndex))
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(item.PdfFilePath) && string.IsNullOrWhiteSpace(item.FileSha256))
                {
                    RegisterLocalFailure(localFailures, item, "pdf_missing", "PDF output is missing.");
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(item.PdfFilePath) && !File.Exists(item.PdfFilePath))
                {
                    RegisterLocalFailure(localFailures, item, "pdf_missing", "PDF file was not found on disk.");
                    continue;
                }

                if (includeNative && item.IncludeNative)
                {
                    if (string.IsNullOrWhiteSpace(item.NativeFilePath))
                    {
                        RegisterLocalFailure(localFailures, item, "native_missing", "Native DWG output is missing.");
                        continue;
                    }

                    if (!File.Exists(item.NativeFilePath))
                    {
                        RegisterLocalFailure(localFailures, item, "native_missing", "Native DWG file was not found on disk.");
                    }
                }
            }

            return localFailures;
        }

        private static Dictionary<int, ExportArtifact> BuildArtifactMap(
            IReadOnlyList<ExportArtifact> artifacts,
            string expectedKind)
        {
            Dictionary<int, ExportArtifact> map = new Dictionary<int, ExportArtifact>();
            if (artifacts == null)
            {
                return map;
            }

            for (int i = 0; i < artifacts.Count; i++)
            {
                ExportArtifact artifact = artifacts[i];
                if (artifact == null)
                {
                    continue;
                }

                if (!string.Equals(artifact.Kind, expectedKind, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                map[artifact.ItemIndex] = artifact;
            }

            return map;
        }

        private static void RegisterLocalFailure(
            Dictionary<int, PublishItemResult> localFailures,
            PublishSheetItem item,
            string errorCode,
            string errorMessage)
        {
            if (localFailures.ContainsKey(item.ItemIndex))
            {
                return;
            }

            localFailures[item.ItemIndex] = new PublishItemResult
            {
                ItemIndex = item.ItemIndex,
                State = "failed",
                DocNumber = item.DocNumber ?? string.Empty,
                AppliedRevision = item.RequestedRevision ?? string.Empty,
                ErrorCode = NormalizeErrorCode(errorCode, "export_failed"),
                ErrorMessage = errorMessage ?? string.Empty,
            };
        }

        private static string NormalizeErrorCode(string value, string fallback)
        {
            return string.IsNullOrWhiteSpace(value) ? fallback : value.Trim();
        }

        private static async Task ExecuteLoginPreflightAsync(
            IApiClient apiClient,
            PublishSheetsCommandRequest request,
            CancellationToken cancellationToken)
        {
            try
            {
                await apiClient
                    .LoginAsync(request.Username, request.Password, cancellationToken)
                    .ConfigureAwait(false);
            }
            catch (TaskCanceledException ex)
            {
                throw new InvalidOperationException("api_unreachable: login timed out before export started.", ex);
            }
            catch (HttpRequestException ex)
            {
                if (IsAuthFailure(ex))
                {
                    throw new InvalidOperationException("auth_failed: login rejected by MDR API.", ex);
                }

                throw new InvalidOperationException("api_unreachable: login endpoint is unreachable.", ex);
            }
        }

        private static bool IsAuthFailure(HttpRequestException ex)
        {
            string message = ex.Message ?? string.Empty;
            return message.IndexOf("401", StringComparison.OrdinalIgnoreCase) >= 0 ||
                message.IndexOf("403", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static PublishBatchResponse MergeWithLocalFailures(
            PublishBuildResult buildResult,
            PublishBatchResponse serverResponse)
        {
            Dictionary<int, PublishItemResult> serverByIndex = new Dictionary<int, PublishItemResult>();
            if (serverResponse != null && serverResponse.Items != null)
            {
                for (int i = 0; i < serverResponse.Items.Count; i++)
                {
                    PublishItemResult item = serverResponse.Items[i];
                    if (item == null)
                    {
                        continue;
                    }

                    serverByIndex[item.ItemIndex] = CloneItemResult(item);
                }
            }

            PublishBatchResponse merged = new PublishBatchResponse
            {
                RunId = string.IsNullOrWhiteSpace(serverResponse?.RunId)
                    ? buildResult.PublishRequest.RunClientId
                    : serverResponse.RunId,
            };

            for (int i = 0; i < buildResult.RequestedItemIndexes.Count; i++)
            {
                int itemIndex = buildResult.RequestedItemIndexes[i];
                if (buildResult.LocalFailures.TryGetValue(itemIndex, out PublishItemResult? localFailure))
                {
                    merged.Items.Add(CloneItemResult(localFailure));
                    continue;
                }

                if (serverByIndex.TryGetValue(itemIndex, out PublishItemResult? serverItem))
                {
                    merged.Items.Add(CloneItemResult(serverItem));
                    continue;
                }

                merged.Items.Add(new PublishItemResult
                {
                    ItemIndex = itemIndex,
                    State = "failed",
                    ErrorCode = "result_missing",
                    ErrorMessage = "No item-level result returned by API.",
                });
            }

            merged.Summary = BuildSummary(merged.Items);
            return merged;
        }

        private static PublishBatchSummary BuildSummary(IReadOnlyList<PublishItemResult> items)
        {
            PublishBatchSummary summary = new PublishBatchSummary
            {
                RequestedCount = items?.Count ?? 0,
            };

            if (items != null)
            {
                for (int i = 0; i < items.Count; i++)
                {
                    PublishItemResult item = items[i];
                    string state = (item?.State ?? string.Empty).Trim();
                    if (string.Equals(state, "completed", StringComparison.OrdinalIgnoreCase))
                    {
                        summary.SuccessCount++;
                        continue;
                    }

                    if (string.Equals(state, "duplicate", StringComparison.OrdinalIgnoreCase))
                    {
                        summary.DuplicateCount++;
                        continue;
                    }

                    summary.FailedCount++;
                }
            }

            if (summary.RequestedCount == 0 || summary.FailedCount >= summary.RequestedCount)
            {
                summary.Status = "failed";
            }
            else if (summary.FailedCount == 0 && summary.DuplicateCount == 0)
            {
                summary.Status = "completed";
            }
            else
            {
                summary.Status = "completed_with_errors";
            }

            return summary;
        }

        private static PublishItemResult CloneItemResult(PublishItemResult item)
        {
            return new PublishItemResult
            {
                ItemIndex = item.ItemIndex,
                State = item.State,
                DocumentId = item.DocumentId,
                DocNumber = item.DocNumber,
                AppliedRevision = item.AppliedRevision,
                PdfFileId = item.PdfFileId,
                NativeFileId = item.NativeFileId,
                ErrorCode = item.ErrorCode,
                ErrorMessage = item.ErrorMessage,
            };
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

        private sealed class PublishBuildResult
        {
            public PublishBatchRequest PublishRequest { get; set; } = new PublishBatchRequest();

            public List<int> RequestedItemIndexes { get; set; } = new List<int>();

            public Dictionary<int, PublishItemResult> LocalFailures { get; set; } = new Dictionary<int, PublishItemResult>();

            public string OutputDirectory { get; set; } = string.Empty;
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

        public int RequestTimeoutSeconds { get; set; } = 120;

        public bool AllowInsecureTls { get; set; }

        public string NativeFormat { get; set; } = "dwg";

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
