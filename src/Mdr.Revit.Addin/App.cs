using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Autodesk.Revit.UI;
using Mdr.Revit.Addin.Commands;
using Mdr.Revit.Addin.Ribbon;
using Mdr.Revit.Core.Models;
using Mdr.Revit.Infra.Config;

namespace Mdr.Revit.Addin
{
    public sealed class App
    {
        private readonly RibbonBuilder _ribbonBuilder;
        private readonly SettingsCommand _settingsCommand;
        private readonly LoginCommand _loginCommand;
        private readonly PublishSheetsCommand _publishSheetsCommand;
        private readonly PushSchedulesCommand _pushSchedulesCommand;
        private readonly SyncSiteLogsCommand _syncSiteLogsCommand;
        private readonly GoogleSyncCommand _googleSyncCommand;
        private readonly SmartNumberingCommand _smartNumberingCommand;
        private readonly CheckUpdatesCommand _checkUpdatesCommand;

        public App()
            : this(
                new RibbonBuilder(),
                new SettingsCommand(),
                new LoginCommand(),
                new PublishSheetsCommand(),
                new PushSchedulesCommand(),
                new SyncSiteLogsCommand(),
                new GoogleSyncCommand(),
                new SmartNumberingCommand(),
                new CheckUpdatesCommand())
        {
        }

        public App(UIDocument uiDocument)
            : this(
                new RibbonBuilder(),
                new SettingsCommand(),
                new LoginCommand(),
                new PublishSheetsCommand(uiDocument),
                new PushSchedulesCommand(uiDocument),
                new SyncSiteLogsCommand(),
                new GoogleSyncCommand(uiDocument),
                new SmartNumberingCommand(uiDocument),
                new CheckUpdatesCommand())
        {
        }

        internal App(
            RibbonBuilder ribbonBuilder,
            SettingsCommand settingsCommand,
            LoginCommand loginCommand,
            PublishSheetsCommand publishSheetsCommand,
            PushSchedulesCommand pushSchedulesCommand,
            SyncSiteLogsCommand syncSiteLogsCommand,
            GoogleSyncCommand googleSyncCommand,
            SmartNumberingCommand smartNumberingCommand,
            CheckUpdatesCommand checkUpdatesCommand)
        {
            _ribbonBuilder = ribbonBuilder ?? throw new ArgumentNullException(nameof(ribbonBuilder));
            _settingsCommand = settingsCommand ?? throw new ArgumentNullException(nameof(settingsCommand));
            _loginCommand = loginCommand ?? throw new ArgumentNullException(nameof(loginCommand));
            _publishSheetsCommand = publishSheetsCommand ?? throw new ArgumentNullException(nameof(publishSheetsCommand));
            _pushSchedulesCommand = pushSchedulesCommand ?? throw new ArgumentNullException(nameof(pushSchedulesCommand));
            _syncSiteLogsCommand = syncSiteLogsCommand ?? throw new ArgumentNullException(nameof(syncSiteLogsCommand));
            _googleSyncCommand = googleSyncCommand ?? throw new ArgumentNullException(nameof(googleSyncCommand));
            _smartNumberingCommand = smartNumberingCommand ?? throw new ArgumentNullException(nameof(smartNumberingCommand));
            _checkUpdatesCommand = checkUpdatesCommand ?? throw new ArgumentNullException(nameof(checkUpdatesCommand));
        }

        public string Name => "MDR Revit Plugin";

        public string ConfigPath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MDR",
            "RevitPlugin",
            "config.json");

        public IReadOnlyList<RibbonCommandDescriptor> BuildRibbon()
        {
            return _ribbonBuilder.Build();
        }

        public PluginConfig LoadConfig()
        {
            PluginConfig config = _settingsCommand.Load(ConfigPath);
            EnsureConfigPersisted(config);
            return config;
        }

        public void SaveConfig(PluginConfig config)
        {
            _settingsCommand.Save(ConfigPath, config);
        }

        public Task<string> LoginAsync(
            string username,
            string password,
            CancellationToken cancellationToken)
        {
            PluginConfig config = LoadConfig();
            LoginCommandRequest request = new LoginCommandRequest
            {
                BaseUrl = config.ApiBaseUrl,
                Username = username,
                Password = password,
                RequestTimeoutSeconds = config.RequestTimeoutSeconds,
                AllowInsecureTls = config.AllowInsecureTls,
            };
            return _loginCommand.ExecuteAsync(request, cancellationToken);
        }

        public Task<PublishSheetsCommandResult> PublishSelectedSheetsAsync(
            PublishFromAppRequest request,
            CancellationToken cancellationToken)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            PluginConfig config = LoadConfig();
            PublishSheetsCommandRequest commandRequest = new PublishSheetsCommandRequest
            {
                BaseUrl = config.ApiBaseUrl,
                Username = request.Username,
                Password = request.Password,
                ProjectCode = ResolveProjectCode(request.ProjectCode, config.ProjectCode),
                ModelGuid = request.ModelGuid ?? string.Empty,
                ModelTitle = request.ModelTitle ?? string.Empty,
                RevitVersion = string.IsNullOrWhiteSpace(request.RevitVersion) ? "2026" : request.RevitVersion,
                PluginVersion = string.IsNullOrWhiteSpace(config.PluginVersion) ? "0.3.0" : config.PluginVersion,
                DefaultStatusCode = string.IsNullOrWhiteSpace(config.DefaultPublishStatusCode)
                    ? "IFA"
                    : config.DefaultPublishStatusCode,
                IncludeNative = request.IncludeNative ?? config.IncludeNativeByDefault,
                RetryFailedItems = request.RetryFailedItems ?? config.RetryFailedItems,
                OutputDirectory = ResolveOutputDirectory(request.OutputDirectory, config.PublishOutputDirectory),
                RequestTimeoutSeconds = config.RequestTimeoutSeconds,
                AllowInsecureTls = config.AllowInsecureTls,
                NativeFormat = ResolvePublishNativeFormat(config),
            };

            foreach (PublishSheetItem item in request.Items)
            {
                commandRequest.Items.Add(item);
            }

            return _publishSheetsCommand.ExecuteAsync(commandRequest, cancellationToken);
        }

        public IReadOnlyList<PublishSheetItem> GetSelectedSheetsForPublish()
        {
            return _publishSheetsCommand.GetSelectedSheets();
        }

        public Task<ScheduleIngestResponse> PushSchedulesAsync(
            PushSchedulesFromAppRequest request,
            CancellationToken cancellationToken)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            PluginConfig config = LoadConfig();
            PushSchedulesCommandRequest commandRequest = new PushSchedulesCommandRequest
            {
                BaseUrl = config.ApiBaseUrl,
                Username = request.Username,
                Password = request.Password,
                ProjectCode = ResolveProjectCode(request.ProjectCode, config.ProjectCode),
                ProfileCode = string.IsNullOrWhiteSpace(request.ProfileCode) ? ScheduleProfiles.Mto : request.ProfileCode,
                ModelGuid = request.ModelGuid ?? string.Empty,
                ViewName = request.ViewName ?? string.Empty,
                SchemaVersion = string.IsNullOrWhiteSpace(request.SchemaVersion) ? "v1" : request.SchemaVersion,
                RequestTimeoutSeconds = config.RequestTimeoutSeconds,
                AllowInsecureTls = config.AllowInsecureTls,
            };

            return _pushSchedulesCommand.ExecuteAsync(commandRequest, cancellationToken);
        }

        public Task<SiteLogApplyResult> SyncSiteLogsAsync(
            SyncSiteLogsFromAppRequest request,
            CancellationToken cancellationToken)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            PluginConfig config = LoadConfig();
            SyncSiteLogsCommandRequest commandRequest = new SyncSiteLogsCommandRequest
            {
                BaseUrl = config.ApiBaseUrl,
                Username = request.Username,
                Password = request.Password,
                ProjectCode = ResolveProjectCode(request.ProjectCode, config.ProjectCode),
                DisciplineCode = request.DisciplineCode ?? string.Empty,
                ClientModelGuid = request.ClientModelGuid ?? string.Empty,
                UpdatedAfterUtc = request.UpdatedAfterUtc,
                Limit = request.Limit <= 0 ? 500 : request.Limit,
                PluginVersion = string.IsNullOrWhiteSpace(config.PluginVersion) ? "0.3.0" : config.PluginVersion,
                RequestTimeoutSeconds = config.RequestTimeoutSeconds,
                AllowInsecureTls = config.AllowInsecureTls,
            };

            return _syncSiteLogsCommand.ExecuteAsync(commandRequest, cancellationToken);
        }

        public Task<GoogleScheduleSyncResult> SyncGoogleSheetsAsync(
            GoogleSyncFromAppRequest request,
            CancellationToken cancellationToken)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            PluginConfig config = LoadConfig();
            GoogleSyncCommandRequest commandRequest = new GoogleSyncCommandRequest
            {
                Direction = string.IsNullOrWhiteSpace(request.Direction)
                    ? GoogleSyncDirections.Export
                    : request.Direction,
                ScheduleName = request.ScheduleName ?? string.Empty,
                SpreadsheetId = string.IsNullOrWhiteSpace(request.SpreadsheetId)
                    ? config.Google.DefaultSpreadsheetId
                    : request.SpreadsheetId,
                WorksheetName = string.IsNullOrWhiteSpace(request.WorksheetName)
                    ? config.Google.DefaultWorksheetName
                    : request.WorksheetName,
                AnchorColumn = string.IsNullOrWhiteSpace(request.AnchorColumn)
                    ? "MDR_UNIQUE_ID"
                    : request.AnchorColumn,
                GoogleClientId = config.Google.ClientId,
                GoogleClientSecret = config.Google.ClientSecret,
                RefreshToken = config.Google.RefreshToken,
                TokenStorePath = ResolveGoogleTokenStorePath(config.Google.TokenStorePath),
                AuthorizeInteractively = request.AuthorizeInteractively,
                PreviewOnly = request.PreviewOnly,
            };

            foreach (GoogleSheetColumnMapping mapping in request.ColumnMappings)
            {
                commandRequest.ColumnMappings.Add(mapping);
            }

            foreach (string column in config.Google.ProtectedSystemColumns)
            {
                commandRequest.ProtectedColumns.Add(column);
            }

            return _googleSyncCommand.ExecuteAsync(commandRequest, cancellationToken);
        }

        public IReadOnlyList<string> GetAvailableSchedulesForGoogleSync()
        {
            return _googleSyncCommand.GetAvailableSchedules();
        }

        public SmartNumberingResult ApplySmartNumbering(SmartNumberingFromAppRequest request)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            PluginConfig config = LoadConfig();
            SmartNumberingRule? rule = ResolveSmartNumberingRule(request, config);
            if (rule == null)
            {
                throw new InvalidOperationException("No smart numbering rule is configured.");
            }

            return _smartNumberingCommand.Execute(new SmartNumberingCommandRequest
            {
                Rule = rule,
                PreviewOnly = request.PreviewOnly,
            });
        }

        public Task<UpdateCheckResult> CheckUpdatesAsync(
            CheckUpdatesFromAppRequest request,
            CancellationToken cancellationToken)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            PluginConfig config = LoadConfig();
            CheckUpdatesCommandRequest commandRequest = new CheckUpdatesCommandRequest
            {
                CurrentVersion = string.IsNullOrWhiteSpace(config.PluginVersion) ? "0.3.0" : config.PluginVersion,
                Channel = string.IsNullOrWhiteSpace(request.Channel) ? config.Updates.Channel : request.Channel,
                GithubRepo = string.IsNullOrWhiteSpace(request.GithubRepo) ? config.Updates.GithubRepo : request.GithubRepo,
                DownloadDirectory = ResolveUpdatesDirectory(request.DownloadDirectory),
                RequireSignature = config.Updates.RequireSignature,
            };

            foreach (string thumbprint in config.Updates.AllowedPublisherThumbprints)
            {
                commandRequest.AllowedPublisherThumbprints.Add(thumbprint);
            }

            return _checkUpdatesCommand.ExecuteAsync(commandRequest, cancellationToken);
        }

        private void EnsureConfigPersisted(PluginConfig config)
        {
            if (!File.Exists(ConfigPath))
            {
                SaveConfig(config);
            }
        }

        private static string ResolveProjectCode(string requestedProjectCode, string defaultProjectCode)
        {
            string value = !string.IsNullOrWhiteSpace(requestedProjectCode)
                ? requestedProjectCode
                : defaultProjectCode;

            if (string.IsNullOrWhiteSpace(value))
            {
                throw new InvalidOperationException("ProjectCode is required.");
            }

            return value;
        }

        private static string ResolveOutputDirectory(string requestedOutputDirectory, string configuredOutputDirectory)
        {
            string value = !string.IsNullOrWhiteSpace(requestedOutputDirectory)
                ? requestedOutputDirectory
                : configuredOutputDirectory;

            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            return ExpandEnvironmentTokens(value);
        }

        private static string ExpandEnvironmentTokens(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return string.Empty;
            }

            return Environment.ExpandEnvironmentVariables(path.Replace("/", "\\"));
        }

        private static string ResolvePublishNativeFormat(PluginConfig config)
        {
            if (config?.Publish != null && !string.IsNullOrWhiteSpace(config.Publish.NativeFormat))
            {
                return config.Publish.NativeFormat;
            }

            return "dwg";
        }

        private static string ResolveGoogleTokenStorePath(string configuredPath)
        {
            if (!string.IsNullOrWhiteSpace(configuredPath))
            {
                return ExpandEnvironmentTokens(configuredPath);
            }

            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "MDR",
                "RevitPlugin",
                "google",
                "token.dat");
        }

        private static string ResolveUpdatesDirectory(string requestedDirectory)
        {
            if (!string.IsNullOrWhiteSpace(requestedDirectory))
            {
                return ExpandEnvironmentTokens(requestedDirectory);
            }

            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "MDR",
                "RevitPlugin",
                "updates");
        }

        private static SmartNumberingRule? ResolveSmartNumberingRule(
            SmartNumberingFromAppRequest request,
            PluginConfig config)
        {
            if (request.Rule != null)
            {
                return request.Rule;
            }

            string selectedRuleId = string.IsNullOrWhiteSpace(request.RuleId)
                ? config.SmartNumbering.DefaultRuleId
                : request.RuleId;
            if (string.IsNullOrWhiteSpace(selectedRuleId))
            {
                return config.SmartNumbering.Rules.Count > 0
                    ? config.SmartNumbering.Rules[0]
                    : null;
            }

            for (int i = 0; i < config.SmartNumbering.Rules.Count; i++)
            {
                SmartNumberingRule rule = config.SmartNumbering.Rules[i];
                if (string.Equals(rule.RuleId, selectedRuleId, StringComparison.OrdinalIgnoreCase))
                {
                    return rule;
                }
            }

            return null;
        }
    }

    public sealed class PublishFromAppRequest
    {
        public string Username { get; set; } = string.Empty;

        public string Password { get; set; } = string.Empty;

        public string ProjectCode { get; set; } = string.Empty;

        public string ModelGuid { get; set; } = string.Empty;

        public string ModelTitle { get; set; } = string.Empty;

        public string RevitVersion { get; set; } = "2026";

        public bool? IncludeNative { get; set; }

        public bool? RetryFailedItems { get; set; }

        public string OutputDirectory { get; set; } = string.Empty;

        public List<PublishSheetItem> Items { get; } = new List<PublishSheetItem>();
    }

    public sealed class PushSchedulesFromAppRequest
    {
        public string Username { get; set; } = string.Empty;

        public string Password { get; set; } = string.Empty;

        public string ProjectCode { get; set; } = string.Empty;

        public string ProfileCode { get; set; } = ScheduleProfiles.Mto;

        public string ModelGuid { get; set; } = string.Empty;

        public string ViewName { get; set; } = string.Empty;

        public string SchemaVersion { get; set; } = "v1";
    }

    public sealed class SyncSiteLogsFromAppRequest
    {
        public string Username { get; set; } = string.Empty;

        public string Password { get; set; } = string.Empty;

        public string ProjectCode { get; set; } = string.Empty;

        public string DisciplineCode { get; set; } = string.Empty;

        public string ClientModelGuid { get; set; } = string.Empty;

        public DateTimeOffset? UpdatedAfterUtc { get; set; }

        public int Limit { get; set; } = 500;
    }

    public sealed class GoogleSyncFromAppRequest
    {
        public string Direction { get; set; } = GoogleSyncDirections.Export;

        public string ScheduleName { get; set; } = string.Empty;

        public string SpreadsheetId { get; set; } = string.Empty;

        public string WorksheetName { get; set; } = string.Empty;

        public string AnchorColumn { get; set; } = "MDR_UNIQUE_ID";

        public bool AuthorizeInteractively { get; set; }

        public bool PreviewOnly { get; set; }

        public List<GoogleSheetColumnMapping> ColumnMappings { get; } = new List<GoogleSheetColumnMapping>();
    }

    public sealed class SmartNumberingFromAppRequest
    {
        public string RuleId { get; set; } = string.Empty;

        public SmartNumberingRule? Rule { get; set; }

        public bool PreviewOnly { get; set; }
    }

    public sealed class CheckUpdatesFromAppRequest
    {
        public string Channel { get; set; } = string.Empty;

        public string GithubRepo { get; set; } = string.Empty;

        public string DownloadDirectory { get; set; } = string.Empty;
    }
}
