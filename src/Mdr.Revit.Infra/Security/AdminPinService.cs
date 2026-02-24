using System;
using System.Security.Cryptography;
using Mdr.Revit.Infra.Config;

namespace Mdr.Revit.Infra.Security
{
    public sealed class AdminPinService
    {
        private const int MinPinLength = 6;
        private const int MaxPinLength = 32;
        private const int SaltBytes = 16;
        private const int HashBytes = 32;
        private static readonly object StateLock = new object();
        private static int _failedAttempts;
        private static DateTimeOffset? _lockedUntilUtc;

        public bool IsPinConfigured(PluginConfig config)
        {
            if (config == null)
            {
                throw new ArgumentNullException(nameof(config));
            }

            if (config.AdminMode == null)
            {
                return false;
            }

            return !string.IsNullOrWhiteSpace(config.AdminMode.PinHash) &&
                   !string.IsNullOrWhiteSpace(config.AdminMode.PinSalt);
        }

        public void ConfigurePin(PluginConfig config, string pin, string confirm)
        {
            if (config == null)
            {
                throw new ArgumentNullException(nameof(config));
            }

            ValidatePin(pin);
            ValidatePin(confirm);

            if (!string.Equals(pin, confirm, StringComparison.Ordinal))
            {
                throw new InvalidOperationException("PIN and confirmation do not match.");
            }

            byte[] salt = new byte[SaltBytes];
            RandomNumberGenerator.Fill(salt);

            int iterations = NormalizeIterations(config.AdminMode?.PinIterations ?? 0);
            byte[] hash = Derive(pin, salt, iterations);

            if (config.AdminMode == null)
            {
                config.AdminMode = new AdminModePluginConfig();
            }

            config.AdminMode.PinSalt = Convert.ToBase64String(salt);
            config.AdminMode.PinHash = Convert.ToBase64String(hash);
            config.AdminMode.PinIterations = iterations;

            ResetFailures();
        }

        public AdminPinVerificationResult VerifyPin(PluginConfig config, string pin)
        {
            if (config == null)
            {
                throw new ArgumentNullException(nameof(config));
            }

            if (config.AdminMode == null)
            {
                return AdminPinVerificationResult.Fail("admin_mode_missing", "Admin mode configuration is missing.");
            }

            if (!config.AdminMode.Enabled)
            {
                return AdminPinVerificationResult.Success();
            }

            if (IsLockedOut(config, out int retryAfterSeconds))
            {
                return AdminPinVerificationResult.LockedOut(retryAfterSeconds);
            }

            if (!IsPinConfigured(config))
            {
                return AdminPinVerificationResult.Fail("pin_not_configured", "Admin PIN is not configured.");
            }

            if (string.IsNullOrWhiteSpace(pin))
            {
                RegisterFailure(config);
                return BuildInvalidPinResult(config);
            }

            byte[] expectedHash;
            byte[] salt;
            try
            {
                expectedHash = Convert.FromBase64String(config.AdminMode.PinHash ?? string.Empty);
                salt = Convert.FromBase64String(config.AdminMode.PinSalt ?? string.Empty);
            }
            catch (FormatException)
            {
                return AdminPinVerificationResult.Fail("pin_storage_invalid", "Stored PIN hash/salt format is invalid.");
            }

            int iterations = NormalizeIterations(config.AdminMode.PinIterations);
            byte[] actualHash = Derive(pin, salt, iterations);
            bool isMatch = expectedHash.Length == actualHash.Length &&
                           CryptographicOperations.FixedTimeEquals(expectedHash, actualHash);
            if (isMatch)
            {
                ResetFailures();
                return AdminPinVerificationResult.Success();
            }

            RegisterFailure(config);
            if (IsLockedOut(config, out retryAfterSeconds))
            {
                return AdminPinVerificationResult.LockedOut(retryAfterSeconds);
            }

            return BuildInvalidPinResult(config);
        }

        public void RegisterFailure(PluginConfig config)
        {
            if (config == null)
            {
                throw new ArgumentNullException(nameof(config));
            }

            int maxAttempts = NormalizeMaxAttempts(config.AdminMode?.MaxAttempts ?? 0);
            int lockoutSeconds = NormalizeLockoutSeconds(config.AdminMode?.LockoutSeconds ?? 0);

            lock (StateLock)
            {
                _failedAttempts++;
                if (_failedAttempts >= maxAttempts)
                {
                    _failedAttempts = 0;
                    _lockedUntilUtc = DateTimeOffset.UtcNow.AddSeconds(lockoutSeconds);
                }
            }
        }

        public void ResetFailures()
        {
            lock (StateLock)
            {
                _failedAttempts = 0;
                _lockedUntilUtc = null;
            }
        }

        public bool IsLockedOut(PluginConfig config, out int retryAfterSeconds)
        {
            _ = config;
            lock (StateLock)
            {
                if (!_lockedUntilUtc.HasValue || DateTimeOffset.UtcNow >= _lockedUntilUtc.Value)
                {
                    _lockedUntilUtc = null;
                    retryAfterSeconds = 0;
                    return false;
                }

                retryAfterSeconds = (int)Math.Ceiling((_lockedUntilUtc.Value - DateTimeOffset.UtcNow).TotalSeconds);
                if (retryAfterSeconds < 0)
                {
                    retryAfterSeconds = 0;
                }

                return true;
            }
        }

        public int GetRemainingAttempts(PluginConfig config)
        {
            if (config == null)
            {
                throw new ArgumentNullException(nameof(config));
            }

            int maxAttempts = NormalizeMaxAttempts(config.AdminMode?.MaxAttempts ?? 0);
            lock (StateLock)
            {
                int remaining = maxAttempts - _failedAttempts;
                return remaining < 0 ? 0 : remaining;
            }
        }

        private AdminPinVerificationResult BuildInvalidPinResult(PluginConfig config)
        {
            return AdminPinVerificationResult.Fail(
                "pin_invalid",
                "PIN is invalid.",
                GetRemainingAttempts(config));
        }

        private static byte[] Derive(string pin, byte[] salt, int iterations)
        {
            return Rfc2898DeriveBytes.Pbkdf2(
                pin,
                salt,
                iterations,
                HashAlgorithmName.SHA256,
                HashBytes);
        }

        private static void ValidatePin(string pin)
        {
            if (string.IsNullOrWhiteSpace(pin))
            {
                throw new InvalidOperationException("PIN is required.");
            }

            if (pin.Length < MinPinLength || pin.Length > MaxPinLength)
            {
                throw new InvalidOperationException("PIN length must be between 6 and 32 characters.");
            }
        }

        private static int NormalizeIterations(int value)
        {
            return value <= 0 ? 120000 : value;
        }

        private static int NormalizeMaxAttempts(int value)
        {
            return value <= 0 ? 5 : value;
        }

        private static int NormalizeLockoutSeconds(int value)
        {
            return value <= 0 ? 60 : value;
        }
    }

    public sealed class AdminPinVerificationResult
    {
        public bool IsSuccess { get; private set; }

        public bool IsLockedOut { get; private set; }

        public string ErrorCode { get; private set; } = string.Empty;

        public string Message { get; private set; } = string.Empty;

        public int RemainingAttempts { get; private set; }

        public int RetryAfterSeconds { get; private set; }

        public static AdminPinVerificationResult Success()
        {
            return new AdminPinVerificationResult
            {
                IsSuccess = true,
            };
        }

        public static AdminPinVerificationResult Fail(string errorCode, string message, int remainingAttempts = 0)
        {
            return new AdminPinVerificationResult
            {
                IsSuccess = false,
                ErrorCode = errorCode ?? string.Empty,
                Message = message ?? string.Empty,
                RemainingAttempts = remainingAttempts < 0 ? 0 : remainingAttempts,
            };
        }

        public static AdminPinVerificationResult LockedOut(int retryAfterSeconds)
        {
            return new AdminPinVerificationResult
            {
                IsSuccess = false,
                IsLockedOut = true,
                ErrorCode = "pin_locked_out",
                Message = "Too many invalid attempts. Try again later.",
                RetryAfterSeconds = retryAfterSeconds < 0 ? 0 : retryAfterSeconds,
            };
        }
    }
}
