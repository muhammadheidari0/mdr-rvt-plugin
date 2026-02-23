using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;
using Mdr.Revit.Core.UseCases;
using Xunit;

namespace Mdr.Revit.Core.Tests
{
    public sealed class CheckForUpdatesUseCaseTests
    {
        [Fact]
        public async Task ExecuteAsync_WhenUpdateAvailable_DownloadsAndPreparesInstall()
        {
            FakeUpdateFeed feed = new FakeUpdateFeed("1.2.0");
            FakeInstaller installer = new FakeInstaller();
            CheckForUpdatesUseCase useCase = new CheckForUpdatesUseCase(feed, installer);

            UpdateCheckResult result = await useCase.ExecuteAsync(
                new UpdateCheckRequest
                {
                    CurrentVersion = "1.1.0",
                    GithubRepo = "owner/repo",
                },
                CancellationToken.None);

            Assert.True(result.IsUpdateAvailable);
            Assert.Equal("1.2.0", result.LatestVersion);
            Assert.Equal(1, installer.DownloadCount);
            Assert.Equal(1, installer.PrepareCount);
            Assert.True(result.Install.IsReady);
        }

        private sealed class FakeUpdateFeed : IUpdateFeedClient
        {
            private readonly string _version;

            public FakeUpdateFeed(string version)
            {
                _version = version;
            }

            public Task<UpdateManifest> GetLatestAsync(UpdateCheckRequest request, CancellationToken cancellationToken)
            {
                _ = request;
                _ = cancellationToken;
                UpdateManifest manifest = new UpdateManifest
                {
                    Version = _version,
                };
                manifest.Assets.Add(new UpdateAsset
                {
                    Name = "Mdr.Revit.Plugin.msi",
                    DownloadUrl = "https://example.test/plugin.msi",
                });
                return Task.FromResult(manifest);
            }
        }

        private sealed class FakeInstaller : IUpdateInstaller
        {
            public int DownloadCount { get; private set; }

            public int PrepareCount { get; private set; }

            public Task<UpdateInstallResult> DownloadAndVerifyAsync(
                UpdateManifest manifest,
                UpdateCheckRequest request,
                CancellationToken cancellationToken)
            {
                _ = manifest;
                _ = request;
                _ = cancellationToken;
                DownloadCount++;
                return Task.FromResult(new UpdateInstallResult
                {
                    IsReady = true,
                    DownloadedFilePath = "C:\\temp\\plugin.msi",
                });
            }

            public Task<UpdateInstallResult> PrepareInstallAsync(
                UpdateInstallResult downloadResult,
                CancellationToken cancellationToken)
            {
                _ = cancellationToken;
                PrepareCount++;
                downloadResult.InstallCommand = "install-update.cmd";
                return Task.FromResult(downloadResult);
            }
        }
    }
}
