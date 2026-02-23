using System;
using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Core.Contracts;

namespace Mdr.Revit.Client.Http
{
    public sealed class AuthClient
    {
        private readonly IApiClient _apiClient;

        public AuthClient(IApiClient apiClient)
        {
            _apiClient = apiClient ?? throw new ArgumentNullException(nameof(apiClient));
        }

        public Task<string> LoginAsync(string username, string password, CancellationToken cancellationToken)
        {
            return _apiClient.LoginAsync(username, password, cancellationToken);
        }
    }
}
