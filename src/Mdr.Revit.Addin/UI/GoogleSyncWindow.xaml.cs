using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Addin.UI
{
    public sealed class GoogleSyncWindow
    {
        private readonly App _app;

        public GoogleSyncWindowViewModel ViewModel { get; }

        public GoogleSyncWindow()
            : this(new App())
        {
        }

        internal GoogleSyncWindow(App app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            ViewModel = new GoogleSyncWindowViewModel();
            ReloadDefaultsFromConfig();
            ReloadAvailableSchedules();
        }

        public Task<GoogleScheduleSyncResult> SyncAsync(
            GoogleSyncFromAppRequest request,
            CancellationToken cancellationToken)
        {
            return _app.SyncGoogleSheetsAsync(request, cancellationToken);
        }

        public Task<GoogleScheduleSyncResult> SyncAsync(CancellationToken cancellationToken)
        {
            return _app.SyncGoogleSheetsAsync(ViewModel.BuildRequest(), cancellationToken);
        }

        public void ReloadAvailableSchedules()
        {
            ViewModel.SetScheduleNames(_app.GetAvailableSchedulesForGoogleSync());
        }

        public void ReloadDefaultsFromConfig()
        {
            var config = _app.LoadConfig();
            ViewModel.SetDefaults(
                config.Google.DefaultSpreadsheetId,
                config.Google.DefaultWorksheetName,
                "MDR_UNIQUE_ID");
            ViewModel.ApplyProtectedColumns(config.Google.ProtectedSystemColumns);
        }

        public bool? ShowDialog()
        {
            Window window = CreateWindow();

            ComboBox directionCombo = new ComboBox
            {
                Width = 220,
                Margin = new Thickness(8, 0, 0, 0),
                ItemsSource = ViewModel.Directions,
                SelectedItem = ViewModel.Direction,
            };

            ComboBox scheduleCombo = new ComboBox
            {
                Width = 320,
                Margin = new Thickness(8, 0, 0, 0),
                ItemsSource = ViewModel.ScheduleNames,
                SelectedItem = ViewModel.SelectedScheduleName,
            };

            TextBox spreadsheetTextBox = new TextBox
            {
                Margin = new Thickness(8, 0, 0, 0),
                MinWidth = 640,
                Text = ViewModel.SpreadsheetId,
            };

            TextBox worksheetTextBox = new TextBox
            {
                Width = 220,
                Margin = new Thickness(8, 0, 0, 0),
                Text = ViewModel.WorksheetName,
            };

            TextBox anchorTextBox = new TextBox
            {
                Width = 320,
                Margin = new Thickness(8, 0, 0, 0),
                Text = ViewModel.AnchorColumn,
            };

            CheckBox previewCheckBox = new CheckBox
            {
                Content = "Preview only",
                Margin = new Thickness(0, 0, 16, 0),
                IsChecked = ViewModel.PreviewOnly,
            };

            CheckBox authorizeCheckBox = new CheckBox
            {
                Content = "Authorize interactively (OAuth)",
                IsChecked = ViewModel.AuthorizeInteractively,
            };

            DataGrid mappingsGrid = BuildMappingsGrid();

            Button refreshButton = new Button
            {
                Width = 130,
                Margin = new Thickness(0, 0, 8, 0),
                Content = "Refresh Schedules",
            };

            Button syncButton = new Button
            {
                Width = 120,
                Margin = new Thickness(0, 0, 8, 0),
                Content = "Sync to Google",
                IsDefault = true,
            };

            Button closeButton = new Button
            {
                Width = 90,
                Content = "Close",
                IsCancel = true,
            };

            refreshButton.Click += (_, _) =>
            {
                ReloadAvailableSchedules();
                scheduleCombo.ItemsSource = null;
                scheduleCombo.ItemsSource = ViewModel.ScheduleNames;
                scheduleCombo.SelectedItem = ViewModel.SelectedScheduleName;
            };

            syncButton.Click += (_, _) =>
            {
                _ = RunSyncAsync(
                    window,
                    directionCombo,
                    scheduleCombo,
                    spreadsheetTextBox,
                    worksheetTextBox,
                    anchorTextBox,
                    previewCheckBox,
                    authorizeCheckBox,
                    refreshButton,
                    syncButton,
                    closeButton);
            };

            closeButton.Click += (_, _) =>
            {
                window.DialogResult = false;
                window.Close();
            };

            Grid root = BuildLayout(
                directionCombo,
                scheduleCombo,
                spreadsheetTextBox,
                worksheetTextBox,
                anchorTextBox,
                previewCheckBox,
                authorizeCheckBox,
                mappingsGrid,
                refreshButton,
                syncButton,
                closeButton);
            window.Content = root;
            return window.ShowDialog();
        }

        private static Window CreateWindow()
        {
            return new Window
            {
                Title = "Google Sheets Sync",
                Width = 920,
                Height = 640,
                MinWidth = 760,
                MinHeight = 520,
                ResizeMode = ResizeMode.CanResize,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
            };
        }

        private static DataGrid BuildMappingsGrid()
        {
            DataGrid grid = new DataGrid
            {
                AutoGenerateColumns = false,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                Margin = new Thickness(0, 0, 0, 12),
            };

            DataGridTextColumn sheetColumn = new DataGridTextColumn
            {
                Header = "Sheet Column",
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                Binding = new Binding(nameof(GoogleSyncMappingItem.SheetColumn)),
            };
            DataGridTextColumn parameterColumn = new DataGridTextColumn
            {
                Header = "Revit Parameter",
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                Binding = new Binding(nameof(GoogleSyncMappingItem.RevitParameter)),
            };
            DataGridCheckBoxColumn editableColumn = new DataGridCheckBoxColumn
            {
                Header = "Editable",
                Width = 120,
                Binding = new Binding(nameof(GoogleSyncMappingItem.IsEditable)),
            };

            grid.Columns.Add(sheetColumn);
            grid.Columns.Add(parameterColumn);
            grid.Columns.Add(editableColumn);

            return grid;
        }

        private Grid BuildLayout(
            ComboBox directionCombo,
            ComboBox scheduleCombo,
            TextBox spreadsheetTextBox,
            TextBox worksheetTextBox,
            TextBox anchorTextBox,
            CheckBox previewCheckBox,
            CheckBox authorizeCheckBox,
            DataGrid mappingsGrid,
            Button refreshButton,
            Button syncButton,
            Button closeButton)
        {
            Grid root = new Grid
            {
                Margin = new Thickness(16),
            };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            TextBlock title = new TextBlock
            {
                Margin = new Thickness(0, 0, 0, 8),
                FontWeight = FontWeights.SemiBold,
                Text = "Sync Revit Schedule <-> Google Sheets",
            };
            root.Children.Add(title);
            Grid.SetRow(title, 0);

            Grid formGrid = new Grid
            {
                Margin = new Thickness(0, 0, 0, 10),
            };
            formGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            formGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            formGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            formGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            formGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            formGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            formGrid.Children.Add(CreateLabeledRow("Direction", directionCombo, 0, 0, 8));
            formGrid.Children.Add(CreateLabeledRow("Schedule", scheduleCombo, 0, 1, 0));
            formGrid.Children.Add(CreateLabeledRow("Spreadsheet ID", spreadsheetTextBox, 1, 0, 0, 2));
            formGrid.Children.Add(CreateLabeledRow("Worksheet", worksheetTextBox, 2, 0, 8));
            formGrid.Children.Add(CreateLabeledRow("Anchor Column", anchorTextBox, 2, 1, 0));

            StackPanel options = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 0, 0, 0),
            };
            options.Children.Add(previewCheckBox);
            options.Children.Add(authorizeCheckBox);
            formGrid.Children.Add(options);
            Grid.SetRow(options, 3);
            Grid.SetColumn(options, 0);
            Grid.SetColumnSpan(options, 2);

            root.Children.Add(formGrid);
            Grid.SetRow(formGrid, 1);

            mappingsGrid.ItemsSource = ViewModel.ColumnMappings;
            root.Children.Add(mappingsGrid);
            Grid.SetRow(mappingsGrid, 2);

            StackPanel buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 0, 0, 0),
            };
            buttons.Children.Add(refreshButton);
            buttons.Children.Add(syncButton);
            buttons.Children.Add(closeButton);
            root.Children.Add(buttons);
            Grid.SetRow(buttons, 3);

            return root;
        }

        private static FrameworkElement CreateLabeledRow(
            string label,
            FrameworkElement input,
            int row,
            int column,
            double rightMargin,
            int columnSpan = 1)
        {
            StackPanel panel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 0, rightMargin, 8),
            };
            panel.Children.Add(new TextBlock
            {
                Width = 110,
                VerticalAlignment = VerticalAlignment.Center,
                Text = label,
            });
            panel.Children.Add(input);
            Grid.SetRow(panel, row);
            Grid.SetColumn(panel, column);
            Grid.SetColumnSpan(panel, columnSpan);
            return panel;
        }

        private void ApplyControlValues(
            ComboBox directionCombo,
            ComboBox scheduleCombo,
            TextBox spreadsheetTextBox,
            TextBox worksheetTextBox,
            TextBox anchorTextBox,
            CheckBox previewCheckBox,
            CheckBox authorizeCheckBox)
        {
            ViewModel.Direction = directionCombo.SelectedItem as string ?? GoogleSyncDirections.Export;
            ViewModel.SelectedScheduleName = scheduleCombo.SelectedItem as string ?? string.Empty;
            ViewModel.SpreadsheetId = spreadsheetTextBox.Text ?? string.Empty;
            ViewModel.WorksheetName = worksheetTextBox.Text ?? string.Empty;
            ViewModel.AnchorColumn = anchorTextBox.Text ?? "MDR_UNIQUE_ID";
            ViewModel.PreviewOnly = previewCheckBox.IsChecked ?? false;
            ViewModel.AuthorizeInteractively = authorizeCheckBox.IsChecked ?? false;
        }

        private async Task RunSyncAsync(
            Window window,
            ComboBox directionCombo,
            ComboBox scheduleCombo,
            TextBox spreadsheetTextBox,
            TextBox worksheetTextBox,
            TextBox anchorTextBox,
            CheckBox previewCheckBox,
            CheckBox authorizeCheckBox,
            Button refreshButton,
            Button syncButton,
            Button closeButton)
        {
            string originalSyncText = syncButton.Content?.ToString() ?? "Sync to Google";
            try
            {
                syncButton.IsEnabled = false;
                refreshButton.IsEnabled = false;
                closeButton.IsEnabled = false;
                syncButton.Content = "Syncing...";

                ApplyControlValues(
                    directionCombo,
                    scheduleCombo,
                    spreadsheetTextBox,
                    worksheetTextBox,
                    anchorTextBox,
                    previewCheckBox,
                    authorizeCheckBox);

                if (string.Equals(ViewModel.Direction, GoogleSyncDirections.Import, StringComparison.OrdinalIgnoreCase))
                {
                    GoogleScheduleSyncResult previewResult = await RunImportAsync(previewOnly: true).ConfigureAwait(true);
                    DiffViewerWindow previewWindow = new DiffViewerWindow();
                    previewWindow.SetDiff(previewResult.DiffResult);

                    bool shouldApply = previewWindow.ShowDialog(allowApply: !ViewModel.PreviewOnly);
                    if (ViewModel.PreviewOnly || !shouldApply)
                    {
                        ShowSyncSummary(previewResult);
                        return;
                    }

                    GoogleScheduleSyncResult applyResult = await RunImportAsync(previewOnly: false).ConfigureAwait(true);
                    ShowSyncSummary(applyResult);
                    window.DialogResult = true;
                    window.Close();
                    return;
                }

                GoogleScheduleSyncResult exportResult = await SyncAsync(CancellationToken.None).ConfigureAwait(true);
                ShowSyncSummary(exportResult);
                window.DialogResult = true;
                window.Close();
            }
            catch (TimeoutException ex)
            {
                MessageBox.Show(
                    ex.Message,
                    "Google OAuth Timeout",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message,
                    "Google Sync Failed",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                syncButton.Content = originalSyncText;
                syncButton.IsEnabled = true;
                refreshButton.IsEnabled = true;
                closeButton.IsEnabled = true;
            }
        }

        private async Task<GoogleScheduleSyncResult> RunImportAsync(bool previewOnly)
        {
            GoogleSyncFromAppRequest request = ViewModel.BuildRequest();
            request.Direction = GoogleSyncDirections.Import;
            request.PreviewOnly = previewOnly;
            return await SyncAsync(request, CancellationToken.None).ConfigureAwait(true);
        }

        private static void ShowSyncSummary(GoogleScheduleSyncResult result)
        {
            string message;
            if (string.Equals(result.Direction, GoogleSyncDirections.Export, StringComparison.OrdinalIgnoreCase))
            {
                message = "Export completed.\n" +
                    "Rows exported: " + result.ExportedRows + "\n" +
                    "Updated range: " + (result.WriteResult.UpdatedRange ?? string.Empty);
            }
            else
            {
                message = "Import completed.\n" +
                    "Changed rows: " + result.DiffResult.ChangedRowsCount + "\n" +
                    "Errors: " + result.DiffResult.ErrorRowsCount + "\n" +
                    "Applied: " + result.ApplyResult.AppliedCount + "\n" +
                    "Failed: " + result.ApplyResult.FailedCount;
            }

            MessageBox.Show(
                message,
                "Google Sync",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
    }
}
