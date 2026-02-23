using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Core.Contracts
{
    public interface IUpdateFeedClient
    {
        Task<UpdateManifest> GetLatestAsync(UpdateCheckRequest request, CancellationToken cancellationToken);
    }
}
