using System;
using System.IO;
using Mdr.Revit.Infra.Config;
using Mdr.Revit.Infra.Security;
using Xunit;

namespace Mdr.Revit.Addin.Tests
{
    public sealed class ExcelImportPasswordServiceTests
    {
        [Fact]
        public void ConfigureAndVerifyPassword_StoresHashOnly()
        {
            PluginConfig config = new PluginConfig();
            ExcelImportPasswordService service = new ExcelImportPasswordService();

            service.ConfigurePassword(config, "secret123", "secret123");

            Assert.True(service.IsPasswordConfigured(config));
            Assert.NotEqual("secret123", config.Excel.ImportPasswordHash);
            Assert.NotEqual("secret123", config.Excel.ImportPasswordSalt);
            Assert.True(service.VerifyPassword(config, "secret123").IsSuccess);
            Assert.False(service.VerifyPassword(config, "wrong123").IsSuccess);
        }

        [Fact]
        public void Load_BackwardCompatibleConfig_AddsExcelDefaults()
        {
            string path = Path.Combine(Path.GetTempPath(), "mdr-config-excel-" + Guid.NewGuid().ToString("N") + ".json");
            try
            {
                File.WriteAllText(path, "{ \"apiBaseUrl\": \"http://127.0.0.1:8000\" }");
                ConfigLoader loader = new ConfigLoader();

                PluginConfig config = loader.Load(path);

                Assert.NotNull(config.Excel);
                Assert.Equal("%LocalAppData%/MDR/RevitPlugin/excel", config.Excel.DefaultDirectory);
                Assert.Equal("MDR_UNIQUE_ID", config.Excel.AnchorColumn);
                Assert.Equal(120000, config.Excel.ImportPasswordIterations);
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
