using System;
using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Addin.UI
{
    public sealed class ScheduleWindow
    {
        private readonly App _app;

        public ScheduleWindow()
            : this(new App())
        {
        }

        internal ScheduleWindow(App app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
        }

        public Task<ScheduleIngestResponse> PushAsync(
            PushSchedulesFromAppRequest request,
            CancellationToken cancellationToken)
        {
            return _app.PushSchedulesAsync(request, cancellationToken);
        }
    }
}
