using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Autodesk.Revit.UI;
using Mdr.Revit.Addin.Security;
using Mdr.Revit.Client.Auth;
using Mdr.Revit.Client.Http;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;
using Mdr.Revit.Core.UseCases;
using Mdr.Revit.Infra.Logging;
using Mdr.Revit.RevitAdapter.Extractors;

namespace Mdr.Revit.Addin.Commands
{
    public sealed class GoogleSyncCommand
    {
        private readonly IRevitScheduleSyncAdapter _revitScheduleSyncAdapter;
        private readonly PluginLogger _logger;

        public GoogleSyncCommand()
            : this(
                new NullScheduleSyncAdapter(),
                new PluginLogger(DefaultLogDirectory()))
        {
        }

        public GoogleSyncCommand(UIDocument uiDocument)
            : this(
                RevitApiExtractors.CreateScheduleSyncAdapter(uiDocument),
                new PluginLogger(DefaultLogDirectory()))
        {
        }

        internal GoogleSyncCommand(IRevitScheduleSyncAdapter revitScheduleSyncAdapter, PluginLogger logger)
        {
            _revitScheduleSyncAdapter = revitScheduleSyncAdapter ?? throw new ArgumentNullException(nameof(revitScheduleSyncAdapter));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        public string Id => "mdr.googleSync";

        public IReadOnlyList<string> GetAvailableSchedules()
        {
            return _revitScheduleSyncAdapter.GetAvailableScheduleNames();
        }

        public async Task<GoogleScheduleSyncResult> ExecuteAsync(
            GoogleSyncCommandRequest request,
            CancellationToken cancellationToken)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            ApplyEmbeddedCredentialsFallback(request);
            ValidateRequest(request);
            _logger.Info("Starting Google Sheets sync direction=" + request.Direction);

            string tokenFile = ResolveTokenFilePath(request.TokenStorePath);
            GoogleTokenStore tokenStore = new GoogleTokenStore(tokenFile);
            await EnsureTokenMaterializedAsync(request, tokenStore, cancellationToken).ConfigureAwait(false);

            using (GoogleDesktopTokenProvider tokenProvider = new GoogleDesktopTokenProvider(
                request.GoogleClientId,
                request.GoogleClientSecret,
                tokenStore))
            using (GoogleSheetsClient googleClient = new GoogleSheetsClient(tokenProvider))
            {
                GoogleSheetSyncProfile profile = BuildProfile(request);
                GoogleScheduleSyncRequest syncRequest = new GoogleScheduleSyncRequest
                {
                    Direction = request.Direction,
                    ScheduleName = request.ScheduleName ?? string.Empty,
                    Profile = profile,
                    ApplyChanges = !request.PreviewOnly,
                };

                if (string.Equals(request.Direction, GoogleSyncDirections.Export, StringComparison.OrdinalIgnoreCase))
                {
                    SyncScheduleToGoogleUseCase useCase = new SyncScheduleToGoogleUseCase(googleClient, _revitScheduleSyncAdapter);
                    return await useCase.ExecuteAsync(syncRequest, cancellationToken).ConfigureAwait(false);
                }

                SyncScheduleFromGoogleUseCase fromUseCase = new SyncScheduleFromGoogleUseCase(googleClient, _revitScheduleSyncAdapter);
                return await fromUseCase.ExecuteAsync(syncRequest, cancellationToken).ConfigureAwait(false);
            }
        }

        private static GoogleSheetSyncProfile BuildProfile(GoogleSyncCommandRequest request)
        {
            GoogleSheetSyncProfile profile = new GoogleSheetSyncProfile
            {
                SpreadsheetId = request.SpreadsheetId,
                WorksheetName = request.WorksheetName,
                AnchorColumn = request.AnchorColumn,
            };

            for (int i = 0; i < request.ColumnMappings.Count; i++)
            {
                profile.ColumnMappings.Add(request.ColumnMappings[i]);
            }

            for (int i = 0; i < request.ProtectedColumns.Count; i++)
            {
                profile.ProtectedColumns.Add(request.ProtectedColumns[i]);
            }

            return profile;
        }

        private static async Task EnsureTokenMaterializedAsync(
            GoogleSyncCommandRequest request,
            GoogleTokenStore tokenStore,
            CancellationToken cancellationToken)
        {
            GoogleOAuthToken token = tokenStore.Load();

            if (!string.IsNullOrWhiteSpace(request.RefreshToken) && string.IsNullOrWhiteSpace(token.RefreshToken))
            {
                token.RefreshToken = request.RefreshToken.Trim();
            }

            if (string.IsNullOrWhiteSpace(token.RefreshToken) && request.AuthorizeInteractively)
            {
                GoogleOAuthDesktopFlow flow = new GoogleOAuthDesktopFlow();
                GoogleOAuthToken interactive = await flow
                    .AuthorizeAsync(request.GoogleClientId, request.GoogleClientSecret, cancellationToken)
                    .ConfigureAwait(false);

                token.AccessToken = interactive.AccessToken;
                token.RefreshToken = interactive.RefreshToken;
                token.ExpiresAtUtc = interactive.ExpiresAtUtc;
            }

            tokenStore.Save(token);
        }

        private static void ValidateRequest(GoogleSyncCommandRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.Direction))
            {
                throw new InvalidOperationException("Direction is required.");
            }

            if (!string.Equals(request.Direction, GoogleSyncDirections.Export, StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(request.Direction, GoogleSyncDirections.Import, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("Direction must be 'export' or 'import'.");
            }

            if (string.IsNullOrWhiteSpace(request.SpreadsheetId))
            {
                throw new InvalidOperationException("SpreadsheetId is required.");
            }

            if (string.IsNullOrWhiteSpace(request.WorksheetName))
            {
                throw new InvalidOperationException("WorksheetName is required.");
            }

            if (string.IsNullOrWhiteSpace(request.GoogleClientId) || string.IsNullOrWhiteSpace(request.GoogleClientSecret))
            {
                throw new InvalidOperationException(
                    "Google OAuth client credentials are required. " +
                    "Set google.clientId/clientSecret in config or embed Resources/Google/credentials.json.");
            }
        }

        private static void ApplyEmbeddedCredentialsFallback(GoogleSyncCommandRequest request)
        {
            bool hasClientId = !string.IsNullOrWhiteSpace(request.GoogleClientId);
            bool hasClientSecret = !string.IsNullOrWhiteSpace(request.GoogleClientSecret);
            if (hasClientId && hasClientSecret)
            {
                return;
            }

            EmbeddedGoogleOAuthCredentials embedded = EmbeddedGoogleOAuthCredentialsProvider.Load();
            if (!embedded.IsConfigured)
            {
                return;
            }

            if (!hasClientId)
            {
                request.GoogleClientId = embedded.ClientId;
            }

            if (!hasClientSecret)
            {
                request.GoogleClientSecret = embedded.ClientSecret;
            }
        }

        private static string ResolveTokenFilePath(string configuredPath)
        {
            if (!string.IsNullOrWhiteSpace(configuredPath))
            {
                return configuredPath;
            }

            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "MDR",
                "RevitPlugin",
                "google",
                "token.dat");
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

    public sealed class GoogleSyncCommandRequest
    {
        public string Direction { get; set; } = GoogleSyncDirections.Export;

        public string ScheduleName { get; set; } = string.Empty;

        public string SpreadsheetId { get; set; } = string.Empty;

        public string WorksheetName { get; set; } = "Sheet1";

        public string AnchorColumn { get; set; } = "MDR_UNIQUE_ID";

        public string GoogleClientId { get; set; } = string.Empty;

        public string GoogleClientSecret { get; set; } = string.Empty;

        public string RefreshToken { get; set; } = string.Empty;

        public string TokenStorePath { get; set; } = string.Empty;

        public bool AuthorizeInteractively { get; set; }

        public bool PreviewOnly { get; set; }

        public List<GoogleSheetColumnMapping> ColumnMappings { get; } = new List<GoogleSheetColumnMapping>();

        public List<string> ProtectedColumns { get; } = new List<string>();
    }

    internal sealed class NullScheduleSyncAdapter : IRevitScheduleSyncAdapter
    {
        public IReadOnlyList<string> GetAvailableScheduleNames()
        {
            return Array.Empty<string>();
        }

        public IReadOnlyList<ScheduleSyncRow> ExtractRows(string scheduleName, GoogleSheetSyncProfile profile)
        {
            _ = scheduleName;
            _ = profile;
            return Array.Empty<ScheduleSyncRow>();
        }

        public ScheduleSyncDiffResult BuildDiff(IReadOnlyList<ScheduleSyncRow> incomingRows, GoogleSheetSyncProfile profile)
        {
            _ = incomingRows;
            _ = profile;
            return new ScheduleSyncDiffResult();
        }

        public ScheduleSyncApplyResult ApplyDiff(ScheduleSyncDiffResult diff, GoogleSheetSyncProfile profile)
        {
            _ = diff;
            _ = profile;
            return new ScheduleSyncApplyResult();
        }
    }
}
