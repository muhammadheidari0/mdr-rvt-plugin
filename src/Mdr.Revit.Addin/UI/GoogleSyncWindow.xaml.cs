using System;
using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Addin.UI
{
    public sealed class GoogleSyncWindow
    {
        private readonly App _app;

        public GoogleSyncWindow()
            : this(new App())
        {
        }

        internal GoogleSyncWindow(App app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
        }

        public Task<GoogleScheduleSyncResult> SyncAsync(
            GoogleSyncFromAppRequest request,
            CancellationToken cancellationToken)
        {
            return _app.SyncGoogleSheetsAsync(request, cancellationToken);
        }
    }
}
