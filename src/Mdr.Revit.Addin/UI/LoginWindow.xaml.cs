using System;
using System.Threading;
using System.Threading.Tasks;

namespace Mdr.Revit.Addin.UI
{
    public sealed class LoginWindow
    {
        private readonly App _app;

        public LoginWindow()
            : this(new App())
        {
        }

        internal LoginWindow(App app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
        }

        public Task<string> LoginAsync(string username, string password, CancellationToken cancellationToken)
        {
            return _app.LoginAsync(username, password, cancellationToken);
        }
    }
}
