using System;
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
    public sealed class PushSchedulesCommand
    {
        private readonly Func<Uri, IApiClient> _apiClientFactory;
        private readonly IRevitExtractor _revitExtractor;
        private readonly PluginLogger _logger;

        public PushSchedulesCommand()
            : this(
                baseAddress => new ApiClient(baseAddress),
                new RevitExtractorAdapter(new SheetExtractor(), new ScheduleExtractor()),
                new PluginLogger(DefaultLogDirectory()))
        {
        }

        public PushSchedulesCommand(UIDocument uiDocument)
            : this(
                baseAddress => new ApiClient(baseAddress),
                RevitApiExtractors.CreateExtractor(uiDocument),
                new PluginLogger(DefaultLogDirectory()))
        {
        }

        internal PushSchedulesCommand(
            Func<Uri, IApiClient> apiClientFactory,
            IRevitExtractor revitExtractor,
            PluginLogger logger)
        {
            _apiClientFactory = apiClientFactory ?? throw new ArgumentNullException(nameof(apiClientFactory));
            _revitExtractor = revitExtractor ?? throw new ArgumentNullException(nameof(revitExtractor));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        public string Id => "mdr.pushSchedules";

        public async Task<ScheduleIngestResponse> ExecuteAsync(
            PushSchedulesCommandRequest request,
            CancellationToken cancellationToken)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            ValidateRequest(request);

            Uri baseAddress = new Uri(request.BaseUrl, UriKind.Absolute);
            IApiClient apiClient = _apiClientFactory(baseAddress);
            PushSchedulesUseCase useCase = new PushSchedulesUseCase(apiClient, _revitExtractor);

            ScheduleIngestRequest ingestRequest = new ScheduleIngestRequest
            {
                ProjectCode = request.ProjectCode,
                ProfileCode = request.ProfileCode,
                ModelGuid = request.ModelGuid,
                ViewName = request.ViewName,
                SchemaVersion = string.IsNullOrWhiteSpace(request.SchemaVersion) ? "v1" : request.SchemaVersion,
                ExtractedAtUtc = DateTimeOffset.UtcNow,
            };

            CorrelationContext.CurrentRunUid = Guid.NewGuid().ToString("N");

            try
            {
                _logger.Info("Starting schedule push.");

                await apiClient
                    .LoginAsync(request.Username, request.Password, cancellationToken)
                    .ConfigureAwait(false);
                _logger.Info("Login succeeded.");

                ScheduleIngestResponse response = await useCase
                    .ExecuteAsync(ingestRequest, cancellationToken)
                    .ConfigureAwait(false);

                CorrelationContext.CurrentRunUid = response.RunId;
                _logger.Info(
                    "Schedule push finished run_id=" + response.RunId +
                    " total=" + response.ValidationSummary.TotalRows +
                    " valid=" + response.ValidationSummary.ValidRows +
                    " invalid=" + response.ValidationSummary.InvalidRows);

                return response;
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

        private static void ValidateRequest(PushSchedulesCommandRequest request)
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

            if (string.IsNullOrWhiteSpace(request.ProfileCode))
            {
                throw new InvalidOperationException("ProfileCode is required.");
            }
        }

        private static string DefaultLogDirectory()
        {
            return System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "MDR",
                "RevitPlugin",
                "logs");
        }
    }

    public sealed class PushSchedulesCommandRequest
    {
        public string BaseUrl { get; set; } = "http://127.0.0.1:8000";

        public string Username { get; set; } = string.Empty;

        public string Password { get; set; } = string.Empty;

        public string ProjectCode { get; set; } = string.Empty;

        public string ProfileCode { get; set; } = ScheduleProfiles.Mto;

        public string ModelGuid { get; set; } = string.Empty;

        public string ViewName { get; set; } = string.Empty;

        public string SchemaVersion { get; set; } = "v1";
    }
}
