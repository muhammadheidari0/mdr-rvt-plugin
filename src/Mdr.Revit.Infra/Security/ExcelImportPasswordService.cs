using System;
using System.Security.Cryptography;
using Mdr.Revit.Infra.Config;

namespace Mdr.Revit.Infra.Security
{
    public sealed class ExcelImportPasswordService
    {
        private const int MinPasswordLength = 6;
        private const int MaxPasswordLength = 64;
        private const int SaltBytes = 16;
        private const int HashBytes = 32;

        public bool IsPasswordConfigured(PluginConfig config)
        {
            if (config == null)
            {
                throw new ArgumentNullException(nameof(config));
            }

            return config.Excel != null &&
                   !string.IsNullOrWhiteSpace(config.Excel.ImportPasswordHash) &&
                   !string.IsNullOrWhiteSpace(config.Excel.ImportPasswordSalt);
        }

        public void ConfigurePassword(PluginConfig config, string password, string confirm)
        {
            if (config == null)
            {
                throw new ArgumentNullException(nameof(config));
            }

            ValidatePassword(password);
            ValidatePassword(confirm);
            if (!string.Equals(password, confirm, StringComparison.Ordinal))
            {
                throw new InvalidOperationException("Excel import password and confirmation do not match.");
            }

            if (config.Excel == null)
            {
                config.Excel = new ExcelPluginConfig();
            }

            byte[] salt = new byte[SaltBytes];
            RandomNumberGenerator.Fill(salt);
            int iterations = NormalizeIterations(config.Excel.ImportPasswordIterations);
            byte[] hash = Derive(password, salt, iterations);

            config.Excel.ImportPasswordSalt = Convert.ToBase64String(salt);
            config.Excel.ImportPasswordHash = Convert.ToBase64String(hash);
            config.Excel.ImportPasswordIterations = iterations;
        }

        public AdminPinVerificationResult VerifyPassword(PluginConfig config, string password)
        {
            if (config == null)
            {
                throw new ArgumentNullException(nameof(config));
            }

            if (config.Excel == null)
            {
                return AdminPinVerificationResult.Fail("excel_config_missing", "Excel configuration is missing.");
            }

            if (!IsPasswordConfigured(config))
            {
                return AdminPinVerificationResult.Fail("excel_password_not_configured", "Excel import password is not configured.");
            }

            if (string.IsNullOrWhiteSpace(password))
            {
                return AdminPinVerificationResult.Fail("excel_password_invalid", "Excel import password is invalid.");
            }

            byte[] expectedHash;
            byte[] salt;
            try
            {
                expectedHash = Convert.FromBase64String(config.Excel.ImportPasswordHash ?? string.Empty);
                salt = Convert.FromBase64String(config.Excel.ImportPasswordSalt ?? string.Empty);
            }
            catch (FormatException)
            {
                return AdminPinVerificationResult.Fail("excel_password_storage_invalid", "Stored Excel import password hash/salt format is invalid.");
            }

            int iterations = NormalizeIterations(config.Excel.ImportPasswordIterations);
            byte[] actualHash = Derive(password, salt, iterations);
            bool isMatch = expectedHash.Length == actualHash.Length &&
                           CryptographicOperations.FixedTimeEquals(expectedHash, actualHash);
            return isMatch
                ? AdminPinVerificationResult.Success()
                : AdminPinVerificationResult.Fail("excel_password_invalid", "Excel import password is invalid.");
        }

        private static void ValidatePassword(string password)
        {
            if (string.IsNullOrWhiteSpace(password))
            {
                throw new InvalidOperationException("Excel import password is required.");
            }

            if (password.Length < MinPasswordLength || password.Length > MaxPasswordLength)
            {
                throw new InvalidOperationException("Excel import password length must be between 6 and 64 characters.");
            }
        }

        private static int NormalizeIterations(int value)
        {
            return value <= 0 ? 120000 : value;
        }

        private static byte[] Derive(string password, byte[] salt, int iterations)
        {
            return Rfc2898DeriveBytes.Pbkdf2(
                password,
                salt,
                iterations,
                HashAlgorithmName.SHA256,
                HashBytes);
        }
    }
}
