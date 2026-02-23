using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Client.Retry;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Client.Http
{
    public sealed class PromptedUpdateInstaller : IUpdateInstaller, IDisposable
    {
        private readonly HttpClient _httpClient;
        private readonly RetryPolicy _retryPolicy;
        private bool _disposed;

        public PromptedUpdateInstaller()
            : this(new HttpClient(), new RetryPolicy())
        {
        }

        public PromptedUpdateInstaller(HttpClient httpClient, RetryPolicy retryPolicy)
        {
            _httpClient = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
            _retryPolicy = retryPolicy ?? throw new ArgumentNullException(nameof(retryPolicy));
        }

        public Task<UpdateInstallResult> DownloadAndVerifyAsync(
            UpdateManifest manifest,
            UpdateCheckRequest request,
            CancellationToken cancellationToken)
        {
            if (manifest == null)
            {
                throw new ArgumentNullException(nameof(manifest));
            }

            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            return _retryPolicy.ExecuteAsync(async token =>
            {
                EnsureNotDisposed();

                UpdateInstallResult result = new UpdateInstallResult();
                UpdateAsset? asset = PickInstallAsset(manifest);
                if (asset == null)
                {
                    result.ValidationErrors.Add("No installable asset (.msi or .zip) was found in release assets.");
                    return result;
                }

                string downloadRoot = string.IsNullOrWhiteSpace(request.DownloadDirectory)
                    ? Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                        "MDR",
                        "RevitPlugin",
                        "updates")
                    : request.DownloadDirectory;
                Directory.CreateDirectory(downloadRoot);

                string filePath = Path.Combine(downloadRoot, asset.Name);
                using (HttpRequestMessage downloadRequest = new HttpRequestMessage(HttpMethod.Get, asset.DownloadUrl))
                {
                    downloadRequest.Headers.UserAgent.ParseAdd("Mdr-Revit-Plugin/1.0");
                    using (HttpResponseMessage response = await _httpClient.SendAsync(downloadRequest, token).ConfigureAwait(false))
                    {
                        if (!response.IsSuccessStatusCode)
                        {
                            result.ValidationErrors.Add(
                                "Failed to download update asset. HTTP " + (int)response.StatusCode + ".");
                            return result;
                        }

                        using (FileStream stream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None))
                        {
                            if (response.Content != null)
                            {
                                await response.Content.CopyToAsync(stream).ConfigureAwait(false);
                            }
                        }
                    }
                }

                result.DownloadedFilePath = filePath;
                await VerifyIntegrityAsync(filePath, asset, request, result, token).ConfigureAwait(false);
                result.IsReady = result.ValidationErrors.Count == 0;
                return result;
            }, cancellationToken);
        }

        public Task<UpdateInstallResult> PrepareInstallAsync(
            UpdateInstallResult downloadResult,
            CancellationToken cancellationToken)
        {
            _ = cancellationToken;
            if (downloadResult == null)
            {
                throw new ArgumentNullException(nameof(downloadResult));
            }

            if (!downloadResult.IsReady || string.IsNullOrWhiteSpace(downloadResult.DownloadedFilePath))
            {
                return Task.FromResult(downloadResult);
            }

            string targetPath = downloadResult.DownloadedFilePath;
            string extension = Path.GetExtension(targetPath);
            string command = extension.Equals(".msi", StringComparison.OrdinalIgnoreCase)
                ? "msiexec /i \"" + targetPath + "\" /passive"
                : "\"" + targetPath + "\"";

            string scriptPath = Path.Combine(
                Path.GetDirectoryName(targetPath) ?? string.Empty,
                "install-update.cmd");
            string script = "@echo off" + Environment.NewLine +
                "timeout /t 2 /nobreak >nul" + Environment.NewLine +
                "start \"\" " + command + Environment.NewLine;
            File.WriteAllText(scriptPath, script, Encoding.ASCII);
            downloadResult.InstallCommand = scriptPath;
            return Task.FromResult(downloadResult);
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _httpClient.Dispose();
            _disposed = true;
        }

        private static UpdateAsset? PickInstallAsset(UpdateManifest manifest)
        {
            if (manifest.Assets == null || manifest.Assets.Count == 0)
            {
                return null;
            }

            UpdateAsset? msi = manifest.Assets.FirstOrDefault(x =>
                x.Name.EndsWith(".msi", StringComparison.OrdinalIgnoreCase));
            if (msi != null)
            {
                return msi;
            }

            return manifest.Assets.FirstOrDefault(x =>
                x.Name.EndsWith(".zip", StringComparison.OrdinalIgnoreCase));
        }

        private async Task VerifyIntegrityAsync(
            string filePath,
            UpdateAsset asset,
            UpdateCheckRequest request,
            UpdateInstallResult result,
            CancellationToken cancellationToken)
        {
            string expectedSha = (asset.Sha256 ?? string.Empty).Trim();
            if (!string.IsNullOrWhiteSpace(expectedSha))
            {
                if (expectedSha.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                {
                    expectedSha = await FetchShaAsync(expectedSha, cancellationToken).ConfigureAwait(false);
                }

                if (!string.IsNullOrWhiteSpace(expectedSha))
                {
                    string actualSha = ComputeSha256(filePath);
                    if (!string.Equals(actualSha, expectedSha, StringComparison.OrdinalIgnoreCase))
                    {
                        result.ValidationErrors.Add("SHA-256 mismatch for downloaded update package.");
                    }
                }
            }

            if (!request.RequireSignature)
            {
                return;
            }

            if (request.AllowedPublisherThumbprints.Count == 0)
            {
                return;
            }

            string certificateThumbprint = ReadSignerThumbprint(filePath);
            if (string.IsNullOrWhiteSpace(certificateThumbprint))
            {
                result.ValidationErrors.Add("Downloaded update package is not authenticode signed.");
                return;
            }

            bool allowed = request.AllowedPublisherThumbprints.Any(x =>
                string.Equals(
                    (x ?? string.Empty).Replace(" ", string.Empty),
                    certificateThumbprint,
                    StringComparison.OrdinalIgnoreCase));
            if (!allowed)
            {
                result.ValidationErrors.Add("Signer certificate thumbprint is not in allowed list.");
            }
        }

        private async Task<string> FetchShaAsync(string shaUrl, CancellationToken cancellationToken)
        {
            using (HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, shaUrl))
            {
                request.Headers.UserAgent.ParseAdd("Mdr-Revit-Plugin/1.0");
                using (HttpResponseMessage response = await _httpClient.SendAsync(request, cancellationToken).ConfigureAwait(false))
                {
                    if (!response.IsSuccessStatusCode)
                    {
                        return string.Empty;
                    }

                    string content = response.Content == null
                        ? string.Empty
                        : await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    if (string.IsNullOrWhiteSpace(content))
                    {
                        return string.Empty;
                    }

                    string[] tokens = content.Trim().Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                    if (tokens.Length == 0)
                    {
                        return string.Empty;
                    }

                    return tokens[0].Trim();
                }
            }
        }

        private static string ComputeSha256(string filePath)
        {
            using (SHA256 sha = SHA256.Create())
            using (FileStream stream = File.OpenRead(filePath))
            {
                byte[] hash = sha.ComputeHash(stream);
                StringBuilder builder = new StringBuilder(hash.Length * 2);
                for (int i = 0; i < hash.Length; i++)
                {
                    builder.Append(hash[i].ToString("x2"));
                }

                return builder.ToString();
            }
        }

        private static string ReadSignerThumbprint(string filePath)
        {
            try
            {
                X509Certificate certificate = X509Certificate.CreateFromSignedFile(filePath);
                X509Certificate2 certificate2 = new X509Certificate2(certificate);
                return (certificate2.Thumbprint ?? string.Empty).Replace(" ", string.Empty);
            }
            catch
            {
                return string.Empty;
            }
        }

        private void EnsureNotDisposed()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(PromptedUpdateInstaller));
            }
        }
    }
}
