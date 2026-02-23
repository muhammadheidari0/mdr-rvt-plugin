using System;
using System.IO;
using System.Text.Json;

namespace Mdr.Revit.Infra.Config
{
    public sealed class ConfigLoader
    {
        private static readonly JsonSerializerOptions JsonSerializerOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true,
            WriteIndented = true,
        };

        public PluginConfig Load(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException("Config path is required.", nameof(path));
            }

            if (!File.Exists(path))
            {
                return PluginConfig.Default;
            }

            string json = File.ReadAllText(path);
            if (string.IsNullOrWhiteSpace(json))
            {
                return PluginConfig.Default;
            }

            PluginConfig? config = JsonSerializer.Deserialize<PluginConfig>(json, JsonSerializerOptions);
            return Normalize(config ?? PluginConfig.Default);
        }

        public void Save(string path, PluginConfig config)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException("Config path is required.", nameof(path));
            }

            if (config == null)
            {
                throw new ArgumentNullException(nameof(config));
            }

            string? directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            PluginConfig normalized = Normalize(config);
            string json = JsonSerializer.Serialize(normalized, JsonSerializerOptions);
            File.WriteAllText(path, json);
        }

        private static PluginConfig Normalize(PluginConfig config)
        {
            if (config == null)
            {
                return PluginConfig.Default;
            }

            if (string.IsNullOrWhiteSpace(config.ApiBaseUrl))
            {
                config.ApiBaseUrl = "http://127.0.0.1:8000";
            }

            if (config.RequestTimeoutSeconds <= 0)
            {
                config.RequestTimeoutSeconds = 120;
            }

            if (string.IsNullOrWhiteSpace(config.PluginVersion))
            {
                config.PluginVersion = "0.3.0";
            }

            if (string.IsNullOrWhiteSpace(config.DefaultPublishStatusCode))
            {
                config.DefaultPublishStatusCode = "IFA";
            }

            if (config.Google == null)
            {
                config.Google = new GooglePluginConfig();
            }

            if (string.IsNullOrWhiteSpace(config.Google.DefaultWorksheetName))
            {
                config.Google.DefaultWorksheetName = "Sheet1";
            }

            if (string.IsNullOrWhiteSpace(config.Google.TokenStorePath))
            {
                config.Google.TokenStorePath = "%LocalAppData%/MDR/RevitPlugin/google/token.dat";
            }

            if (config.Updates == null)
            {
                config.Updates = new UpdatesPluginConfig();
            }

            if (string.IsNullOrWhiteSpace(config.Updates.Channel))
            {
                config.Updates.Channel = "stable";
            }

            if (config.Updates.CheckIntervalHours <= 0)
            {
                config.Updates.CheckIntervalHours = 24;
            }

            if (config.SmartNumbering == null)
            {
                config.SmartNumbering = new SmartNumberingPluginConfig();
            }

            return config;
        }
    }
}
