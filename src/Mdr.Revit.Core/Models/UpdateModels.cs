using System;
using System.Collections.Generic;

namespace Mdr.Revit.Core.Models
{
    public sealed class UpdateCheckRequest
    {
        public string CurrentVersion { get; set; } = string.Empty;

        public string Channel { get; set; } = "stable";

        public string GithubRepo { get; set; } = string.Empty;

        public bool RequireSignature { get; set; } = true;

        public string DownloadDirectory { get; set; } = string.Empty;

        public List<string> AllowedPublisherThumbprints { get; } = new List<string>();
    }

    public sealed class UpdateManifest
    {
        public string Version { get; set; } = string.Empty;

        public DateTimeOffset? PublishedAtUtc { get; set; }

        public string ReleaseNotes { get; set; } = string.Empty;

        public List<UpdateAsset> Assets { get; } = new List<UpdateAsset>();
    }

    public sealed class UpdateAsset
    {
        public string Name { get; set; } = string.Empty;

        public string DownloadUrl { get; set; } = string.Empty;

        public string Sha256 { get; set; } = string.Empty;

        public string Signature { get; set; } = string.Empty;

        public long SizeBytes { get; set; }
    }

    public sealed class UpdateInstallResult
    {
        public bool IsReady { get; set; }

        public string DownloadedFilePath { get; set; } = string.Empty;

        public string InstallCommand { get; set; } = string.Empty;

        public List<string> ValidationErrors { get; } = new List<string>();
    }

    public sealed class UpdateCheckResult
    {
        public string CurrentVersion { get; set; } = string.Empty;

        public string LatestVersion { get; set; } = string.Empty;

        public bool IsUpdateAvailable { get; set; }

        public bool PromptRequired { get; set; } = true;

        public UpdateManifest Manifest { get; set; } = new UpdateManifest();

        public UpdateInstallResult Install { get; set; } = new UpdateInstallResult();
    }
}
