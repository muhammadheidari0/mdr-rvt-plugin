using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Addin.Commands;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Addin.UI
{
    public sealed class PublishWindow
    {
        private readonly App _app;

        public PublishWindowViewModel ViewModel { get; }

        public PublishWindow()
            : this(new App())
        {
        }

        internal PublishWindow(App app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            ViewModel = new PublishWindowViewModel();
            ReloadAvailableSheets();
        }

        public Task<PublishSheetsCommandResult> PublishAsync(
            PublishFromAppRequest request,
            CancellationToken cancellationToken)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            if (!request.IncludeNative.HasValue)
            {
                request.IncludeNative = ViewModel.IncludeNative;
            }

            if (!request.RetryFailedItems.HasValue)
            {
                request.RetryFailedItems = ViewModel.RetryFailedItems;
            }

            if (request.Items.Count == 0)
            {
                IReadOnlyList<PublishSheetItem> selectedItems = ViewModel.BuildSelectedItems();
                for (int i = 0; i < selectedItems.Count; i++)
                {
                    request.Items.Add(selectedItems[i]);
                }
            }

            return _app.PublishSelectedSheetsAsync(request, cancellationToken);
        }

        public void ReloadAvailableSheets()
        {
            ViewModel.SetSheets(_app.GetSelectedSheetsForPublish());
        }
    }
}
