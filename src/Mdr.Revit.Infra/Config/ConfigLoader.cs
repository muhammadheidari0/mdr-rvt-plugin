using System;
using System.IO;
using System.Text.Json;
using Mdr.Revit.Core.Models;

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
                config.PluginVersion = "0.4.0";
            }

            if (string.IsNullOrWhiteSpace(config.DefaultPublishStatusCode))
            {
                config.DefaultPublishStatusCode = "IFA";
            }

            if (config.Publish == null)
            {
                config.Publish = new PublishPluginConfig();
            }

            if (string.IsNullOrWhiteSpace(config.Publish.NativeFormat))
            {
                config.Publish.NativeFormat = "dwg";
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

            if (config.Excel == null)
            {
                config.Excel = new ExcelPluginConfig();
            }

            if (string.IsNullOrWhiteSpace(config.Excel.DefaultDirectory))
            {
                config.Excel.DefaultDirectory = "%LocalAppData%/MDR/RevitPlugin/excel";
            }

            if (string.IsNullOrWhiteSpace(config.Excel.AnchorColumn))
            {
                config.Excel.AnchorColumn = "MDR_UNIQUE_ID";
            }

            if (config.Excel.ImportPasswordIterations <= 0)
            {
                config.Excel.ImportPasswordIterations = 120000;
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

            if (config.AdminMode == null)
            {
                config.AdminMode = new AdminModePluginConfig();
            }

            if (config.AdminMode.PinIterations <= 0)
            {
                config.AdminMode.PinIterations = 120000;
            }

            if (config.AdminMode.MaxAttempts <= 0)
            {
                config.AdminMode.MaxAttempts = 5;
            }

            if (config.AdminMode.LockoutSeconds <= 0)
            {
                config.AdminMode.LockoutSeconds = 60;
            }

            if (config.SmartNumbering == null)
            {
                config.SmartNumbering = new SmartNumberingPluginConfig();
            }

            EnsureSmartNumberingDefaults(config.SmartNumbering);

            return config;
        }

        private static void EnsureSmartNumberingDefaults(SmartNumberingPluginConfig smartNumbering)
        {
            if (smartNumbering.Rules.Count == 0)
            {
                SmartNumberingRule arcaRule = new SmartNumberingRule
                {
                    RuleId = "arca-serial",
                    Mode = SmartNumberingModes.Arca,
                    CategoryBuiltInName = "OST_Walls",
                    Formula = string.Empty,
                    SequenceWidth = 3,
                    StartAt = 1,
                };
                arcaRule.Targets.Add("Serial No");
                smartNumbering.Rules.Add(arcaRule);

                SmartNumberingRule formulaRule = new SmartNumberingRule
                {
                    RuleId = "formula-default",
                    Mode = SmartNumberingModes.Formula,
                    Formula = "{Block}{Level}-{CategoryCode}{SubcategoryCode}{Sequence:5}",
                    SequenceWidth = 5,
                    StartAt = 1,
                };
                formulaRule.Targets.Add("Serial No");
                formulaRule.Targets.Add("Type Mark");
                smartNumbering.Rules.Add(formulaRule);
            }

            if (string.IsNullOrWhiteSpace(smartNumbering.DefaultRuleId))
            {
                smartNumbering.DefaultRuleId = smartNumbering.Rules[0].RuleId;
            }

            for (int i = 0; i < smartNumbering.Rules.Count; i++)
            {
                SmartNumberingRule rule = smartNumbering.Rules[i];
                if (string.IsNullOrWhiteSpace(rule.Mode))
                {
                    rule.Mode = SmartNumberingModes.Formula;
                }

                if (string.Equals(rule.Mode, SmartNumberingModes.Arca, StringComparison.OrdinalIgnoreCase) &&
                    string.IsNullOrWhiteSpace(rule.CategoryBuiltInName))
                {
                    rule.CategoryBuiltInName = "OST_Walls";
                }
            }
        }
    }
}
