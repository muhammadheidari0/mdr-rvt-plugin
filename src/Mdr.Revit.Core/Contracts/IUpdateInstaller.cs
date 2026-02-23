using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Core.Contracts
{
    public interface IUpdateInstaller
    {
        Task<UpdateInstallResult> DownloadAndVerifyAsync(
            UpdateManifest manifest,
            UpdateCheckRequest request,
            CancellationToken cancellationToken);

        Task<UpdateInstallResult> PrepareInstallAsync(
            UpdateInstallResult downloadResult,
            CancellationToken cancellationToken);
    }
}
