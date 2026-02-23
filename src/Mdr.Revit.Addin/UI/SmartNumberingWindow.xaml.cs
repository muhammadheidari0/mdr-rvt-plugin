using System;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Addin.UI
{
    public sealed class SmartNumberingWindow
    {
        private readonly App _app;

        public SmartNumberingWindow()
            : this(new App())
        {
        }

        internal SmartNumberingWindow(App app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
        }

        public SmartNumberingResult Apply(SmartNumberingFromAppRequest request)
        {
            return _app.ApplySmartNumbering(request);
        }
    }
}
