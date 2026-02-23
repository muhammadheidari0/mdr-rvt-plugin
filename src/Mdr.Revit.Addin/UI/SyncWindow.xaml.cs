using System;
using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Addin.UI
{
    public sealed class SyncWindow
    {
        private readonly App _app;

        public SyncWindow()
            : this(new App())
        {
        }

        internal SyncWindow(App app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
        }

        public Task<SiteLogApplyResult> SyncAsync(
            SyncSiteLogsFromAppRequest request,
            CancellationToken cancellationToken)
        {
            return _app.SyncSiteLogsAsync(request, cancellationToken);
        }
    }
}
