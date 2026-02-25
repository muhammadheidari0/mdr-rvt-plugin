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

        public IReadOnlyList<GoogleSheetColumnMapping> GetScheduleColumnMappings(string scheduleName)
        {
            return _revitScheduleSyncAdapter.GetScheduleColumnMappings(scheduleName ?? string.Empty);
        }

        public async Task<GoogleScheduleSyncResult> ExecuteAsync(
            GoogleSyncCommandRequest request,
            CancellationToken cancellationToken)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            NormalizeRequest(request);
            ApplyEmbeddedCredentialsFallback(request);
            ValidateRequest(request);
            _logger.Info("Starting Google Sheets sync direction=" + request.Direction);

            string tokenFile = ResolveTokenFilePath(request.TokenStorePath);
            GoogleTokenStore tokenStore = new GoogleTokenStore(tokenFile);
            await EnsureTokenMaterializedAsync(request, tokenStore, cancellationToken);

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
                    GoogleScheduleSyncResult result = await useCase.ExecuteAsync(syncRequest, cancellationToken);
                    _logger.Info(
                        "Google Sheets export completed rows=" + result.ExportedRows +
                        " skipped=" + result.SkippedRows);

                    foreach (KeyValuePair<string, int> reason in result.SkippedByReason)
                    {
                        _logger.Info("Google Sheets export skipped reason=" + reason.Key + " count=" + reason.Value);
                    }

                    for (int i = 0; i < result.Warnings.Count; i++)
                    {
                        _logger.Info("Google Sheets export warning code=" + result.Warnings[i]);
                    }

                    return result;
                }

                SyncScheduleFromGoogleUseCase fromUseCase = new SyncScheduleFromGoogleUseCase(googleClient, _revitScheduleSyncAdapter);
                GoogleScheduleSyncResult importResult = await fromUseCase.ExecuteAsync(syncRequest, cancellationToken);
                _logger.Info(
                    "Google Sheets import completed changed=" + importResult.DiffResult.ChangedRowsCount +
                    " errors=" + importResult.DiffResult.ErrorRowsCount +
                    " applied=" + importResult.ApplyResult.AppliedCount);
                return importResult;
            }
        }

        private static GoogleSheetSyncProfile BuildProfile(GoogleSyncCommandRequest request)
        {
            HashSet<string> protectedColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "MDR_UNIQUE_ID",
                "MDR_ELEMENT_ID",
            };

            for (int i = 0; i < request.ProtectedColumns.Count; i++)
            {
                string value = (request.ProtectedColumns[i] ?? string.Empty).Trim();
                if (!string.IsNullOrWhiteSpace(value))
                {
                    protectedColumns.Add(value);
                }
            }

            GoogleSheetSyncProfile profile = new GoogleSheetSyncProfile
            {
                SpreadsheetId = request.SpreadsheetId,
                WorksheetName = request.WorksheetName,
                AnchorColumn = request.AnchorColumn,
            };

            HashSet<string> mappedColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < request.ColumnMappings.Count; i++)
            {
                GoogleSheetColumnMapping? mapping = request.ColumnMappings[i];
                if (mapping == null)
                {
                    continue;
                }

                string sheetColumn = (mapping.SheetColumn ?? string.Empty).Trim();
                string revitParameter = (mapping.RevitParameter ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(sheetColumn) || string.IsNullOrWhiteSpace(revitParameter))
                {
                    continue;
                }

                GoogleSheetColumnMapping normalized = new GoogleSheetColumnMapping
                {
                    SheetColumn = sheetColumn,
                    RevitParameter = revitParameter,
                    IsEditable = mapping.IsEditable && !protectedColumns.Contains(sheetColumn),
                };

                if (mappedColumns.Contains(sheetColumn))
                {
                    for (int m = 0; m < profile.ColumnMappings.Count; m++)
                    {
                        if (string.Equals(profile.ColumnMappings[m].SheetColumn, sheetColumn, StringComparison.OrdinalIgnoreCase))
                        {
                            profile.ColumnMappings[m] = normalized;
                            break;
                        }
                    }
                }
                else
                {
                    profile.ColumnMappings.Add(normalized);
                    mappedColumns.Add(sheetColumn);
                }
            }

            foreach (string systemColumn in new[] { "MDR_UNIQUE_ID", "MDR_ELEMENT_ID" })
            {
                if (mappedColumns.Contains(systemColumn))
                {
                    for (int i = 0; i < profile.ColumnMappings.Count; i++)
                    {
                        if (string.Equals(profile.ColumnMappings[i].SheetColumn, systemColumn, StringComparison.OrdinalIgnoreCase))
                        {
                            profile.ColumnMappings[i].IsEditable = false;
                            break;
                        }
                    }
                }
                else
                {
                    profile.ColumnMappings.Add(new GoogleSheetColumnMapping
                    {
                        SheetColumn = systemColumn,
                        RevitParameter = systemColumn,
                        IsEditable = false,
                    });
                }
            }

            foreach (string protectedColumn in protectedColumns)
            {
                profile.ProtectedColumns.Add(protectedColumn);
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
                    .AuthorizeAsync(request.GoogleClientId, request.GoogleClientSecret, cancellationToken);

                token.AccessToken = interactive.AccessToken;
                token.RefreshToken = interactive.RefreshToken;
                token.ExpiresAtUtc = interactive.ExpiresAtUtc;
            }

            tokenStore.Save(token);
        }

        private static void NormalizeRequest(GoogleSyncCommandRequest request)
        {
            request.SpreadsheetId = NormalizeSpreadsheetId(request.SpreadsheetId);
            request.WorksheetName = (request.WorksheetName ?? string.Empty).Trim();
            request.AnchorColumn = string.IsNullOrWhiteSpace(request.AnchorColumn)
                ? "MDR_UNIQUE_ID"
                : request.AnchorColumn.Trim();
        }

        private static string NormalizeSpreadsheetId(string value)
        {
            string input = (value ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(input))
            {
                return string.Empty;
            }

            const string marker = "/spreadsheets/d/";
            int markerIndex = input.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
            if (markerIndex < 0)
            {
                return input;
            }

            string idPart = input.Substring(markerIndex + marker.Length);
            int endIndex = idPart.IndexOfAny(new[] { '/', '?', '#', '&' });
            if (endIndex >= 0)
            {
                idPart = idPart.Substring(0, endIndex);
            }

            return idPart.Trim();
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

        public IReadOnlyList<GoogleSheetColumnMapping> GetScheduleColumnMappings(string scheduleName)
        {
            _ = scheduleName;
            return Array.Empty<GoogleSheetColumnMapping>();
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
