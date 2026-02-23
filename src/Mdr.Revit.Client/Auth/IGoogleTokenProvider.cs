using System.Threading;
using System.Threading.Tasks;

namespace Mdr.Revit.Client.Auth
{
    public interface IGoogleTokenProvider
    {
        Task<string> GetAccessTokenAsync(CancellationToken cancellationToken);
    }
}
