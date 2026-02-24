using Mdr.Revit.Infra.Config;
using Mdr.Revit.Infra.Security;
using Xunit;

namespace Mdr.Revit.Addin.Tests
{
    [Collection("AdminModeSerial")]
    public sealed class AdminPinServiceTests
    {
        [Fact]
        public void ConfigurePin_StoresHashSalt_AndVerifyPasses()
        {
            PluginConfig config = new PluginConfig();
            AdminPinService service = new AdminPinService();
            service.ResetFailures();

            service.ConfigurePin(config, "123456", "123456");
            AdminPinVerificationResult verify = service.VerifyPin(config, "123456");

            Assert.True(service.IsPinConfigured(config));
            Assert.False(string.IsNullOrWhiteSpace(config.AdminMode.PinHash));
            Assert.False(string.IsNullOrWhiteSpace(config.AdminMode.PinSalt));
            Assert.True(verify.IsSuccess);
            Assert.False(verify.IsLockedOut);
        }

        [Fact]
        public void VerifyPin_ReachesMaxAttempts_TriggersLockout()
        {
            PluginConfig config = new PluginConfig();
            config.AdminMode.MaxAttempts = 2;
            config.AdminMode.LockoutSeconds = 60;
            AdminPinService service = new AdminPinService();
            service.ResetFailures();
            service.ConfigurePin(config, "123456", "123456");

            AdminPinVerificationResult first = service.VerifyPin(config, "999999");
            AdminPinVerificationResult second = service.VerifyPin(config, "999999");

            Assert.False(first.IsSuccess);
            Assert.Equal("pin_invalid", first.ErrorCode);
            Assert.Equal(1, first.RemainingAttempts);

            Assert.False(second.IsSuccess);
            Assert.True(second.IsLockedOut);
            Assert.Equal("pin_locked_out", second.ErrorCode);
            Assert.True(second.RetryAfterSeconds > 0);
        }
    }
}
