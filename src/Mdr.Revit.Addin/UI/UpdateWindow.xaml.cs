using System;
using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Addin.UI
{
    public sealed class UpdateWindow
    {
        private readonly App _app;

        public UpdateWindow()
            : this(new App())
        {
        }

        internal UpdateWindow(App app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
        }

        public Task<UpdateCheckResult> CheckAsync(
            CheckUpdatesFromAppRequest request,
            CancellationToken cancellationToken)
        {
            return _app.CheckUpdatesAsync(request, cancellationToken);
        }
    }
}
