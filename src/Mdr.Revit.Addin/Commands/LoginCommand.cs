using System;
using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Core.Contracts;

namespace Mdr.Revit.Addin.Commands
{
    public sealed class LoginCommand
    {
        private readonly Func<ApiClientFactoryOptions, IApiClient> _apiClientFactory;

        public LoginCommand()
            : this(ApiClientFactory.Create)
        {
        }

        internal LoginCommand(Func<ApiClientFactoryOptions, IApiClient> apiClientFactory)
        {
            _apiClientFactory = apiClientFactory ?? throw new ArgumentNullException(nameof(apiClientFactory));
        }

        public string Id => "mdr.login";

        public async Task<string> ExecuteAsync(
            LoginCommandRequest request,
            CancellationToken cancellationToken)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            if (string.IsNullOrWhiteSpace(request.BaseUrl))
            {
                throw new InvalidOperationException("BaseUrl is required.");
            }

            if (string.IsNullOrWhiteSpace(request.Username))
            {
                throw new InvalidOperationException("Username is required.");
            }

            if (string.IsNullOrWhiteSpace(request.Password))
            {
                throw new InvalidOperationException("Password is required.");
            }

            IApiClient apiClient = _apiClientFactory(new ApiClientFactoryOptions
            {
                BaseAddress = new Uri(request.BaseUrl, UriKind.Absolute),
                RequestTimeoutSeconds = request.RequestTimeoutSeconds,
                AllowInsecureTls = request.AllowInsecureTls,
            });
            try
            {
                return await apiClient
                    .LoginAsync(request.Username, request.Password, cancellationToken)
                    .ConfigureAwait(false);
            }
            finally
            {
                if (apiClient is IDisposable disposable)
                {
                    disposable.Dispose();
                }
            }
        }
    }

    public sealed class LoginCommandRequest
    {
        public string BaseUrl { get; set; } = "http://127.0.0.1:8000";

        public string Username { get; set; } = string.Empty;

        public string Password { get; set; } = string.Empty;

        public int RequestTimeoutSeconds { get; set; } = 120;

        public bool AllowInsecureTls { get; set; }
    }
}
