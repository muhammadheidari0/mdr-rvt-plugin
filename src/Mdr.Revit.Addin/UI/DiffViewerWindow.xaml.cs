using System;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Addin.UI
{
    public sealed class DiffViewerWindow
    {
        public ScheduleSyncDiffResult Diff { get; private set; } = new ScheduleSyncDiffResult();

        public void SetDiff(ScheduleSyncDiffResult diff)
        {
            Diff = diff ?? throw new ArgumentNullException(nameof(diff));
        }
    }
}
