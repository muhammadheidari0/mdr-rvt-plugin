using System;
using System.IO;
using Mdr.Revit.Infra.Config;
using Xunit;

namespace Mdr.Revit.Addin.Tests
{
    [Collection("AdminModeSerial")]
    public sealed class ConfigAdminModeTests
    {
        [Fact]
        public void Load_BackwardCompatibleConfig_AddsAdminModeDefaults()
        {
            string path = Path.Combine(Path.GetTempPath(), "mdr-config-admin-" + Guid.NewGuid().ToString("N") + ".json");
            try
            {
                File.WriteAllText(path, "{ \"apiBaseUrl\": \"http://127.0.0.1:8000\", \"publish\": { \"nativeFormat\": \"dwg\" } }");
                ConfigLoader loader = new ConfigLoader();

                PluginConfig config = loader.Load(path);

                Assert.NotNull(config.AdminMode);
                Assert.True(config.AdminMode.Enabled);
                Assert.Equal(120000, config.AdminMode.PinIterations);
                Assert.Equal(5, config.AdminMode.MaxAttempts);
                Assert.Equal(60, config.AdminMode.LockoutSeconds);
            }
            finally
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
        }

        [Fact]
        public void Load_InvalidAdminModeBounds_NormalizesToSafeDefaults()
        {
            string path = Path.Combine(Path.GetTempPath(), "mdr-config-admin-" + Guid.NewGuid().ToString("N") + ".json");
            try
            {
                File.WriteAllText(
                    path,
                    "{ \"apiBaseUrl\": \"http://127.0.0.1:8000\", \"adminMode\": { \"enabled\": true, \"pinIterations\": 0, \"maxAttempts\": 0, \"lockoutSeconds\": 0 } }");
                ConfigLoader loader = new ConfigLoader();

                PluginConfig config = loader.Load(path);

                Assert.Equal(120000, config.AdminMode.PinIterations);
                Assert.Equal(5, config.AdminMode.MaxAttempts);
                Assert.Equal(60, config.AdminMode.LockoutSeconds);
            }
            finally
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
        }
    }
}
