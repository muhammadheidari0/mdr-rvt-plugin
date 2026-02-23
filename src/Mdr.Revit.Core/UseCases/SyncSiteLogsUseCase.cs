using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;
using Mdr.Revit.Core.Validation;

namespace Mdr.Revit.Core.UseCases
{
    public sealed class SyncSiteLogsUseCase
    {
        private readonly IApiClient _apiClient;
        private readonly IRevitWriter _revitWriter;

        public SyncSiteLogsUseCase(IApiClient apiClient, IRevitWriter revitWriter)
        {
            _apiClient = apiClient ?? throw new ArgumentNullException(nameof(apiClient));
            _revitWriter = revitWriter ?? throw new ArgumentNullException(nameof(revitWriter));
        }

        public async Task<SiteLogApplyResult> ExecuteAsync(
            SiteLogManifestRequest manifestRequest,
            string pluginVersion,
            CancellationToken cancellationToken)
        {
            BusinessRules.EnsureManifestRequestIsValid(manifestRequest);

            SiteLogManifestResponse manifest = await _apiClient
                .GetSiteLogManifestAsync(manifestRequest, cancellationToken)
                .ConfigureAwait(false);

            if (manifest.Changes.Count == 0)
            {
                return SiteLogApplyResult.Empty(manifest.RunId);
            }

            SiteLogPullRequest pullRequest = BuildPullRequest(manifestRequest, pluginVersion, manifest.Changes);

            SiteLogPullResponse pullResponse = await _apiClient
                .PullSiteLogRowsAsync(pullRequest, cancellationToken)
                .ConfigureAwait(false);

            SiteLogApplyResult applyResult = _revitWriter.ApplySiteLogRows(pullResponse);
            applyResult.RunId = manifest.RunId;

            SiteLogAckRequest ack = new SiteLogAckRequest
            {
                RunId = manifest.RunId,
                AppliedCount = applyResult.AppliedCount,
                FailedCount = applyResult.FailedCount,
            };

            foreach (SiteLogApplyError error in applyResult.Errors)
            {
                ack.Errors.Add(error);
            }

            await _apiClient
                .AckSiteLogSyncAsync(ack, cancellationToken)
                .ConfigureAwait(false);

            return applyResult;
        }

        private static SiteLogPullRequest BuildPullRequest(
            SiteLogManifestRequest manifestRequest,
            string pluginVersion,
            IEnumerable<SiteLogManifestChange> changes)
        {
            SiteLogPullRequest request = new SiteLogPullRequest
            {
                ProjectCode = manifestRequest.ProjectCode,
                DisciplineCode = manifestRequest.DisciplineCode,
                ClientModelGuid = manifestRequest.ClientModelGuid,
                PluginVersion = pluginVersion ?? string.Empty,
            };

            foreach (long logId in changes.Select(x => x.LogId).Distinct())
            {
                request.LogIds.Add(logId);
            }

            return request;
        }
    }
}
