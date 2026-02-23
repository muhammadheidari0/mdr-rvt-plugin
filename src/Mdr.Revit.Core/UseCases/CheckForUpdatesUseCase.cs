using System;
using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;
using Mdr.Revit.Core.Validation;

namespace Mdr.Revit.Core.UseCases
{
    public sealed class CheckForUpdatesUseCase
    {
        private readonly IUpdateFeedClient _updateFeedClient;
        private readonly IUpdateInstaller _updateInstaller;

        public CheckForUpdatesUseCase(IUpdateFeedClient updateFeedClient, IUpdateInstaller updateInstaller)
        {
            _updateFeedClient = updateFeedClient ?? throw new ArgumentNullException(nameof(updateFeedClient));
            _updateInstaller = updateInstaller ?? throw new ArgumentNullException(nameof(updateInstaller));
        }

        public async Task<UpdateCheckResult> ExecuteAsync(
            UpdateCheckRequest request,
            CancellationToken cancellationToken)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            UpdateManifest latest = await _updateFeedClient
                .GetLatestAsync(request, cancellationToken)
                .ConfigureAwait(false);

            UpdateCheckResult result = new UpdateCheckResult
            {
                CurrentVersion = request.CurrentVersion ?? string.Empty,
                LatestVersion = latest.Version ?? string.Empty,
                Manifest = latest,
            };

            result.IsUpdateAvailable = SemanticVersionComparer.IsGreater(
                candidateVersion: result.LatestVersion,
                currentVersion: result.CurrentVersion);

            if (!result.IsUpdateAvailable)
            {
                return result;
            }

            UpdateInstallResult download = await _updateInstaller
                .DownloadAndVerifyAsync(latest, request, cancellationToken)
                .ConfigureAwait(false);
            result.Install = await _updateInstaller
                .PrepareInstallAsync(download, cancellationToken)
                .ConfigureAwait(false);
            result.PromptRequired = true;
            return result;
        }
    }
}
