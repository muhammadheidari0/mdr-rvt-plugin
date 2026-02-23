using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Client.Http;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;
using Mdr.Revit.Core.UseCases;
using Mdr.Revit.Infra.Logging;

namespace Mdr.Revit.Addin.Commands
{
    public sealed class CheckUpdatesCommand
    {
        private readonly IUpdateFeedClient _updateFeedClient;
        private readonly IUpdateInstaller _updateInstaller;
        private readonly PluginLogger _logger;

        public CheckUpdatesCommand()
            : this(
                new GitHubReleaseClient(),
                new PromptedUpdateInstaller(),
                new PluginLogger(DefaultLogDirectory()))
        {
        }

        internal CheckUpdatesCommand(
            IUpdateFeedClient updateFeedClient,
            IUpdateInstaller updateInstaller,
            PluginLogger logger)
        {
            _updateFeedClient = updateFeedClient ?? throw new ArgumentNullException(nameof(updateFeedClient));
            _updateInstaller = updateInstaller ?? throw new ArgumentNullException(nameof(updateInstaller));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        public string Id => "mdr.checkUpdates";

        public Task<UpdateCheckResult> ExecuteAsync(
            CheckUpdatesCommandRequest request,
            CancellationToken cancellationToken)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            if (string.IsNullOrWhiteSpace(request.CurrentVersion))
            {
                throw new InvalidOperationException("CurrentVersion is required.");
            }

            if (string.IsNullOrWhiteSpace(request.GithubRepo))
            {
                throw new InvalidOperationException("GithubRepo is required.");
            }

            _logger.Info("Checking updates from GitHub repo=" + request.GithubRepo);
            CheckForUpdatesUseCase useCase = new CheckForUpdatesUseCase(_updateFeedClient, _updateInstaller);
            UpdateCheckRequest checkRequest = new UpdateCheckRequest
            {
                CurrentVersion = request.CurrentVersion,
                Channel = string.IsNullOrWhiteSpace(request.Channel) ? "stable" : request.Channel,
                GithubRepo = request.GithubRepo,
                DownloadDirectory = request.DownloadDirectory,
                RequireSignature = request.RequireSignature,
            };

            for (int i = 0; i < request.AllowedPublisherThumbprints.Count; i++)
            {
                checkRequest.AllowedPublisherThumbprints.Add(request.AllowedPublisherThumbprints[i]);
            }

            return useCase.ExecuteAsync(checkRequest, cancellationToken);
        }

        private static string DefaultLogDirectory()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "MDR",
                "RevitPlugin",
                "logs");
        }
    }

    public sealed class CheckUpdatesCommandRequest
    {
        public string CurrentVersion { get; set; } = string.Empty;

        public string Channel { get; set; } = "stable";

        public string GithubRepo { get; set; } = string.Empty;

        public string DownloadDirectory { get; set; } = string.Empty;

        public bool RequireSignature { get; set; } = true;

        public List<string> AllowedPublisherThumbprints { get; } = new List<string>();
    }
}
