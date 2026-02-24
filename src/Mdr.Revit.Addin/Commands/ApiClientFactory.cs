using System;
using System.Net.Http;
using Mdr.Revit.Client.Auth;
using Mdr.Revit.Client.Http;
using Mdr.Revit.Client.Retry;
using Mdr.Revit.Core.Contracts;

namespace Mdr.Revit.Addin.Commands
{
    public sealed class ApiClientFactoryOptions
    {
        public Uri BaseAddress { get; set; } = new Uri("http://127.0.0.1:8000");

        public int RequestTimeoutSeconds { get; set; } = 120;

        public bool AllowInsecureTls { get; set; }
    }

    public static class ApiClientFactory
    {
        public static IApiClient Create(ApiClientFactoryOptions options)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            HttpMessageHandler handler = CreateMessageHandler(options.AllowInsecureTls);
            HttpClient httpClient = new HttpClient(handler)
            {
                BaseAddress = options.BaseAddress,
                Timeout = TimeSpan.FromSeconds(NormalizeTimeout(options.RequestTimeoutSeconds)),
            };

            return new ApiClient(httpClient, new TokenStore(), new RetryPolicy());
        }

        private static HttpMessageHandler CreateMessageHandler(bool allowInsecureTls)
        {
            HttpClientHandler handler = new HttpClientHandler();
            if (allowInsecureTls)
            {
                handler.ServerCertificateCustomValidationCallback =
                    HttpClientHandler.DangerousAcceptAnyServerCertificateValidator;
            }

            return handler;
        }

        private static int NormalizeTimeout(int value)
        {
            if (value <= 0)
            {
                return 120;
            }

            return value;
        }
    }
}
