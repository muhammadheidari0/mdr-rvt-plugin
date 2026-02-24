using System;
using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;
using Mdr.Revit.Core.UseCases;
using Mdr.Revit.Infra.Logging;
using Mdr.Revit.RevitAdapter.Writers;

namespace Mdr.Revit.Addin.Commands
{
    public sealed class SyncSiteLogsCommand
    {
        private readonly Func<ApiClientFactoryOptions, IApiClient> _apiClientFactory;
        private readonly IRevitWriter _revitWriter;
        private readonly PluginLogger _logger;

        public SyncSiteLogsCommand()
            : this(
                ApiClientFactory.Create,
                new SiteLogsScheduleWriter(),
                new PluginLogger(DefaultLogDirectory()))
        {
        }

        internal SyncSiteLogsCommand(
            Func<ApiClientFactoryOptions, IApiClient> apiClientFactory,
            IRevitWriter revitWriter,
            PluginLogger logger)
        {
            _apiClientFactory = apiClientFactory ?? throw new ArgumentNullException(nameof(apiClientFactory));
            _revitWriter = revitWriter ?? throw new ArgumentNullException(nameof(revitWriter));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        public string Id => "mdr.syncSiteLogs";

        public async Task<SiteLogApplyResult> ExecuteAsync(
            SyncSiteLogsCommandRequest request,
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
            SyncSiteLogsUseCase useCase = new SyncSiteLogsUseCase(apiClient, _revitWriter);

            SiteLogManifestRequest manifestRequest = new SiteLogManifestRequest
            {
                ProjectCode = request.ProjectCode,
                DisciplineCode = request.DisciplineCode,
                UpdatedAfterUtc = request.UpdatedAfterUtc,
                Limit = request.Limit,
                ClientModelGuid = request.ClientModelGuid,
            };

            CorrelationContext.CurrentRunUid = Guid.NewGuid().ToString("N");

            try
            {
                _logger.Info("Starting site-log sync.");

                await apiClient
                    .LoginAsync(request.Username, request.Password, cancellationToken)
                    .ConfigureAwait(false);
                _logger.Info("Login succeeded.");

                SiteLogApplyResult result = await useCase
                    .ExecuteAsync(manifestRequest, request.PluginVersion ?? string.Empty, cancellationToken)
                    .ConfigureAwait(false);

                CorrelationContext.CurrentRunUid = result.RunId;
                _logger.Info(
                    "Site-log sync finished run_id=" + result.RunId +
                    " applied=" + result.AppliedCount +
                    " failed=" + result.FailedCount);

                return result;
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

        private static void ValidateRequest(SyncSiteLogsCommandRequest request)
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

            if (string.IsNullOrWhiteSpace(request.ClientModelGuid))
            {
                throw new InvalidOperationException("ClientModelGuid is required.");
            }

            if (request.Limit <= 0 || request.Limit > 10000)
            {
                throw new InvalidOperationException("Limit must be between 1 and 10000.");
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

    public sealed class SyncSiteLogsCommandRequest
    {
        public string BaseUrl { get; set; } = "http://127.0.0.1:8000";

        public string Username { get; set; } = string.Empty;

        public string Password { get; set; } = string.Empty;

        public string ProjectCode { get; set; } = string.Empty;

        public string DisciplineCode { get; set; } = string.Empty;

        public string ClientModelGuid { get; set; } = string.Empty;

        public DateTimeOffset? UpdatedAfterUtc { get; set; }

        public int Limit { get; set; } = 500;

        public string PluginVersion { get; set; } = string.Empty;

        public int RequestTimeoutSeconds { get; set; } = 120;

        public bool AllowInsecureTls { get; set; }
    }
}
