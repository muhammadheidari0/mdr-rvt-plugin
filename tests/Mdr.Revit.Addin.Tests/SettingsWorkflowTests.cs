using System;
using Mdr.Revit.Addin.UI;
using Mdr.Revit.Infra.Config;
using Mdr.Revit.Infra.Security;
using Xunit;

namespace Mdr.Revit.Addin.Tests
{
    [Collection("AdminModeSerial")]
    public sealed class SettingsWorkflowTests
    {
        [Fact]
        public void OpenSettings_FirstRunSetup_ConfiguresPinAndSavesSettings()
        {
            PluginConfig config = new PluginConfig();
            config.AdminMode.PinHash = string.Empty;
            config.AdminMode.PinSalt = string.Empty;
            AdminPinService pinService = new AdminPinService();
            pinService.ResetFailures();
            int saveCount = 0;
            int setupPromptCount = 0;
            int settingsPromptCount = 0;

            SettingsAccessWorkflow workflow = new SettingsAccessWorkflow(
                () => config,
                _ => saveCount++,
                pinService,
                mode =>
                {
                    if (mode != AdminPinDialogMode.Setup)
                    {
                        throw new InvalidOperationException("Unlock prompt is not expected during first-run setup.");
                    }

                    setupPromptCount++;
                    return new AdminPinDialogResult
                    {
                        Accepted = true,
                        Pin = "123456",
                        ConfirmPin = "123456",
                    };
                },
                _ =>
                {
                    settingsPromptCount++;
                    return new SettingsDialogResult
                    {
                        Accepted = true,
                        ApiBaseUrl = "https://mdr.local",
                        NativeFormat = "dwg",
                    };
                },
                (_, _, _) => { });

            bool updated = workflow.OpenSettings();

            Assert.True(updated);
            Assert.True(pinService.IsPinConfigured(config));
            Assert.Equal("https://mdr.local", config.ApiBaseUrl);
            Assert.Equal("dwg", config.Publish.NativeFormat);
            Assert.Equal(1, setupPromptCount);
            Assert.Equal(1, settingsPromptCount);
            Assert.Equal(2, saveCount);
        }

        [Fact]
        public void OpenSettings_UnlockFailureThenCancel_DoesNotSaveSettings()
        {
            PluginConfig config = new PluginConfig();
            AdminPinService pinService = new AdminPinService();
            pinService.ResetFailures();
            pinService.ConfigurePin(config, "123456", "123456");
            int saveCount = 0;
            int promptCount = 0;
            int settingsPromptCount = 0;

            SettingsAccessWorkflow workflow = new SettingsAccessWorkflow(
                () => config,
                _ => saveCount++,
                pinService,
                _ =>
                {
                    promptCount++;
                    if (promptCount == 1)
                    {
                        return new AdminPinDialogResult
                        {
                            Accepted = true,
                            Pin = "000000",
                        };
                    }

                    return new AdminPinDialogResult
                    {
                        Accepted = false,
                    };
                },
                _ =>
                {
                    settingsPromptCount++;
                    return new SettingsDialogResult
                    {
                        Accepted = true,
                        ApiBaseUrl = "https://mdr.local",
                        NativeFormat = "dwg",
                    };
                },
                (_, _, _) => { });

            bool updated = workflow.OpenSettings();

            Assert.False(updated);
            Assert.Equal(0, settingsPromptCount);
            Assert.Equal(0, saveCount);
            Assert.Equal(2, promptCount);
        }

        [Theory]
        [InlineData("https://mdr.local", "dwg", true)]
        [InlineData("http://127.0.0.1:8000", "DWG", true)]
        [InlineData("ftp://mdr.local", "dwg", false)]
        [InlineData("https://mdr.local", "rvt", false)]
        [InlineData("not-a-url", "dwg", false)]
        public void SettingsWindow_Validation_WorksAsExpected(string url, string nativeFormat, bool isValid)
        {
            bool ok = SettingsWindow.TryValidateValues(url, nativeFormat, out string errorMessage);

            Assert.Equal(isValid, ok);
            if (isValid)
            {
                Assert.Equal(string.Empty, errorMessage);
            }
            else
            {
                Assert.False(string.IsNullOrWhiteSpace(errorMessage));
            }
        }

        [Fact]
        public void OpenSettings_InvalidSettingsResult_IsRejectedAndNotSaved()
        {
            PluginConfig config = new PluginConfig();
            AdminPinService pinService = new AdminPinService();
            pinService.ResetFailures();
            pinService.ConfigurePin(config, "123456", "123456");
            int saveCount = 0;

            SettingsAccessWorkflow workflow = new SettingsAccessWorkflow(
                () => config,
                _ => saveCount++,
                pinService,
                _ => new AdminPinDialogResult
                {
                    Accepted = true,
                    Pin = "123456",
                },
                _ => new SettingsDialogResult
                {
                    Accepted = true,
                    ApiBaseUrl = "invalid-url",
                    NativeFormat = "dwg",
                },
                (_, _, _) => { });

            bool updated = workflow.OpenSettings();

            Assert.False(updated);
            Assert.Equal(0, saveCount);
        }
    }
}
