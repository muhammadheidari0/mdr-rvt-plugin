using Mdr.Revit.Addin.Security;
using Xunit;

namespace Mdr.Revit.Addin.Tests
{
    public sealed class EmbeddedGoogleOAuthCredentialsProviderTests
    {
        [Fact]
        public void TryParse_WithInstalledPayload_ReturnsCredentials()
        {
            const string json = @"
{
  ""installed"": {
    ""client_id"": ""abc.apps.googleusercontent.com"",
    ""client_secret"": ""secret-value""
  }
}";

            bool ok = EmbeddedGoogleOAuthCredentialsProvider.TryParse(json, out EmbeddedGoogleOAuthCredentials credentials);

            Assert.True(ok);
            Assert.True(credentials.IsConfigured);
            Assert.Equal("abc.apps.googleusercontent.com", credentials.ClientId);
            Assert.Equal("secret-value", credentials.ClientSecret);
        }

        [Fact]
        public void TryParse_WithWebPayload_ReturnsCredentials()
        {
            const string json = @"
{
  ""web"": {
    ""client_id"": ""web-client.apps.googleusercontent.com"",
    ""client_secret"": ""web-secret""
  }
}";

            bool ok = EmbeddedGoogleOAuthCredentialsProvider.TryParse(json, out EmbeddedGoogleOAuthCredentials credentials);

            Assert.True(ok);
            Assert.Equal("web-client.apps.googleusercontent.com", credentials.ClientId);
            Assert.Equal("web-secret", credentials.ClientSecret);
        }

        [Fact]
        public void TryParse_WithTemplatePlaceholder_ReturnsFalse()
        {
            const string json = @"
{
  ""installed"": {
    ""client_id"": ""YOUR_GOOGLE_CLIENT_ID.apps.googleusercontent.com"",
    ""client_secret"": ""YOUR_GOOGLE_CLIENT_SECRET""
  }
}";

            bool ok = EmbeddedGoogleOAuthCredentialsProvider.TryParse(json, out EmbeddedGoogleOAuthCredentials credentials);

            Assert.False(ok);
            Assert.False(credentials.IsConfigured);
        }
    }
}
