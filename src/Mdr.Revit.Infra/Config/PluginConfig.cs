using System.Collections.Generic;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Infra.Config
{
    public sealed class PluginConfig
    {
        public string ApiBaseUrl { get; set; } = "http://127.0.0.1:8000";

        // Backward-compatible alias for older config files.
        public string BaseUrl
        {
            get
            {
                return ApiBaseUrl;
            }

            set
            {
                ApiBaseUrl = value;
            }
        }

        public string ProjectCode { get; set; } = string.Empty;

        public string PluginVersion { get; set; } = "0.3.6";

        public int RequestTimeoutSeconds { get; set; } = 120;

        public bool AllowInsecureTls { get; set; }

        public string PublishOutputDirectory { get; set; } = string.Empty;

        public string DefaultPublishStatusCode { get; set; } = "IFA";

        public bool IncludeNativeByDefault { get; set; }

        public bool RetryFailedItems { get; set; } = true;

        public PublishPluginConfig Publish { get; set; } = new PublishPluginConfig();

        public GooglePluginConfig Google { get; set; } = new GooglePluginConfig();

        public UpdatesPluginConfig Updates { get; set; } = new UpdatesPluginConfig();

        public AdminModePluginConfig AdminMode { get; set; } = new AdminModePluginConfig();

        public SmartNumberingPluginConfig SmartNumbering { get; set; } = new SmartNumberingPluginConfig();

        public static PluginConfig Default => new PluginConfig();
    }

    public sealed class PublishPluginConfig
    {
        public string NativeFormat { get; set; } = "dwg";
    }

    public sealed class GooglePluginConfig
    {
        public string ClientId { get; set; } = string.Empty;

        public string ClientSecret { get; set; } = string.Empty;

        public string RefreshToken { get; set; } = string.Empty;

        public string TokenStorePath { get; set; } = "%LocalAppData%/MDR/RevitPlugin/google/token.dat";

        public string DefaultSpreadsheetId { get; set; } = string.Empty;

        public string DefaultWorksheetName { get; set; } = "Sheet1";

        public List<string> ProtectedSystemColumns { get; } = new List<string>
        {
            "MDR_UNIQUE_ID",
            "MDR_ELEMENT_ID",
        };
    }

    public sealed class UpdatesPluginConfig
    {
        public bool Enabled { get; set; } = true;

        public string Channel { get; set; } = "stable";

        public string GithubRepo { get; set; } = string.Empty;

        public bool CheckOnStartup { get; set; } = true;

        public int CheckIntervalHours { get; set; } = 24;

        public bool RequireSignature { get; set; } = true;

        public List<string> AllowedPublisherThumbprints { get; } = new List<string>();
    }

    public sealed class AdminModePluginConfig
    {
        public bool Enabled { get; set; } = true;

        public string PinHash { get; set; } = string.Empty;

        public string PinSalt { get; set; } = string.Empty;

        public int PinIterations { get; set; } = 120000;

        public int MaxAttempts { get; set; } = 5;

        public int LockoutSeconds { get; set; } = 60;
    }

    public sealed class SmartNumberingPluginConfig
    {
        public string DefaultRuleId { get; set; } = string.Empty;

        public List<SmartNumberingRule> Rules { get; } = new List<SmartNumberingRule>();

        public List<string> ProtectedReadOnlyParams { get; } = new List<string>
        {
            "MDR_UNIQUE_ID",
            "MDR_ELEMENT_ID",
        };
    }
}
