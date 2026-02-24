using System;
using System.Windows;
using Mdr.Revit.Infra.Config;
using Mdr.Revit.Infra.Security;

namespace Mdr.Revit.Addin.UI
{
    public enum AdminPinDialogMode
    {
        Setup = 0,
        Unlock = 1,
    }

    public sealed class AdminPinDialogResult
    {
        public bool Accepted { get; set; }

        public string Pin { get; set; } = string.Empty;

        public string ConfirmPin { get; set; } = string.Empty;
    }

    public sealed class SettingsDialogResult
    {
        public bool Accepted { get; set; }

        public string ApiBaseUrl { get; set; } = string.Empty;

        public string NativeFormat { get; set; } = "dwg";
    }

    public sealed class SettingsAccessWorkflow
    {
        private readonly Func<PluginConfig> _loadConfig;
        private readonly Action<PluginConfig> _saveConfig;
        private readonly AdminPinService _pinService;
        private readonly Func<AdminPinDialogMode, AdminPinDialogResult> _pinPrompt;
        private readonly Func<PluginConfig, SettingsDialogResult> _settingsPrompt;
        private readonly Action<string, string, MessageBoxImage> _showMessage;

        public SettingsAccessWorkflow(App app)
            : this(
                BuildLoadConfigAccessor(app),
                BuildSaveConfigAccessor(app),
                new AdminPinService(),
                ShowPinDialog,
                ShowSettingsDialog,
                ShowMessage)
        {
        }

        internal SettingsAccessWorkflow(
            Func<PluginConfig> loadConfig,
            Action<PluginConfig> saveConfig,
            AdminPinService pinService,
            Func<AdminPinDialogMode, AdminPinDialogResult> pinPrompt,
            Func<PluginConfig, SettingsDialogResult> settingsPrompt,
            Action<string, string, MessageBoxImage>? showMessage = null)
        {
            _loadConfig = loadConfig ?? throw new ArgumentNullException(nameof(loadConfig));
            _saveConfig = saveConfig ?? throw new ArgumentNullException(nameof(saveConfig));
            _pinService = pinService ?? throw new ArgumentNullException(nameof(pinService));
            _pinPrompt = pinPrompt ?? throw new ArgumentNullException(nameof(pinPrompt));
            _settingsPrompt = settingsPrompt ?? throw new ArgumentNullException(nameof(settingsPrompt));
            _showMessage = showMessage ?? ShowMessage;
        }

        public bool OpenSettings()
        {
            PluginConfig config = _loadConfig();
            if (config == null)
            {
                throw new InvalidOperationException("Plugin config was not loaded.");
            }

            if (config.AdminMode != null && config.AdminMode.Enabled)
            {
                if (!_pinService.IsPinConfigured(config))
                {
                    if (!RunSetup(config))
                    {
                        return false;
                    }
                }
                else if (!RunUnlock(config))
                {
                    return false;
                }
            }

            SettingsDialogResult settingsResult = _settingsPrompt(config);
            if (!settingsResult.Accepted)
            {
                return false;
            }

            if (!SettingsWindow.TryValidateValues(
                    settingsResult.ApiBaseUrl ?? string.Empty,
                    settingsResult.NativeFormat ?? string.Empty,
                    out string validationError))
            {
                _showMessage(validationError, "Invalid Settings", MessageBoxImage.Warning);
                return false;
            }

            config.ApiBaseUrl = settingsResult.ApiBaseUrl ?? string.Empty;
            if (config.Publish == null)
            {
                config.Publish = new PublishPluginConfig();
            }

            config.Publish.NativeFormat = settingsResult.NativeFormat ?? "dwg";
            _saveConfig(config);
            _showMessage("Settings saved successfully.", "Settings", MessageBoxImage.Information);
            return true;
        }

        private bool RunSetup(PluginConfig config)
        {
            while (true)
            {
                AdminPinDialogResult setup = _pinPrompt(AdminPinDialogMode.Setup);
                if (!setup.Accepted)
                {
                    return false;
                }

                try
                {
                    _pinService.ConfigurePin(config, setup.Pin, setup.ConfirmPin);
                    _saveConfig(config);
                    return true;
                }
                catch (Exception ex)
                {
                    _showMessage(ex.Message, "Set Admin PIN", MessageBoxImage.Warning);
                }
            }
        }

        private bool RunUnlock(PluginConfig config)
        {
            while (true)
            {
                if (_pinService.IsLockedOut(config, out int retryAfterSeconds))
                {
                    _showMessage(
                        "Too many invalid attempts. Try again in " + retryAfterSeconds + " seconds.",
                        "Settings Locked",
                        MessageBoxImage.Warning);
                    return false;
                }

                AdminPinDialogResult unlock = _pinPrompt(AdminPinDialogMode.Unlock);
                if (!unlock.Accepted)
                {
                    return false;
                }

                AdminPinVerificationResult verification = _pinService.VerifyPin(config, unlock.Pin);
                if (verification.IsSuccess)
                {
                    return true;
                }

                if (verification.IsLockedOut)
                {
                    _showMessage(
                        "Too many invalid attempts. Try again in " + verification.RetryAfterSeconds + " seconds.",
                        "Settings Locked",
                        MessageBoxImage.Warning);
                    return false;
                }

                string message = string.IsNullOrWhiteSpace(verification.Message)
                    ? "PIN is invalid."
                    : verification.Message;
                if (verification.RemainingAttempts > 0)
                {
                    message += " Remaining attempts: " + verification.RemainingAttempts + ".";
                }

                _showMessage(message, "Invalid PIN", MessageBoxImage.Warning);
            }
        }

        private static AdminPinDialogResult ShowPinDialog(AdminPinDialogMode mode)
        {
            AdminPinDialog dialog = new AdminPinDialog(mode);
            bool accepted = dialog.ShowDialog() == true;
            return new AdminPinDialogResult
            {
                Accepted = accepted,
                Pin = dialog.Pin,
                ConfirmPin = dialog.ConfirmPin,
            };
        }

        private static SettingsDialogResult ShowSettingsDialog(PluginConfig config)
        {
            SettingsWindow dialog = new SettingsWindow(config);
            bool accepted = dialog.ShowDialog() == true;
            return new SettingsDialogResult
            {
                Accepted = accepted,
                ApiBaseUrl = dialog.ApiBaseUrl,
                NativeFormat = dialog.NativeFormat,
            };
        }

        private static void ShowMessage(string message, string title, MessageBoxImage image)
        {
            MessageBox.Show(
                message ?? string.Empty,
                title ?? "Settings",
                MessageBoxButton.OK,
                image);
        }

        private static Func<PluginConfig> BuildLoadConfigAccessor(App app)
        {
            if (app == null)
            {
                throw new ArgumentNullException(nameof(app));
            }

            return app.LoadConfig;
        }

        private static Action<PluginConfig> BuildSaveConfigAccessor(App app)
        {
            if (app == null)
            {
                throw new ArgumentNullException(nameof(app));
            }

            return app.SaveConfig;
        }
    }
}
