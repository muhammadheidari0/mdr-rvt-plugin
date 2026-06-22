using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using Microsoft.Win32;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Addin.UI
{
    public sealed class ExcelSyncWindow
    {
        private readonly App _app;
        private IReadOnlyList<string> _protectedColumns = Array.Empty<string>();

        public ExcelSyncWindowViewModel ViewModel { get; }

        public ExcelSyncWindow()
            : this(new App())
        {
        }

        internal ExcelSyncWindow(App app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            ViewModel = new ExcelSyncWindowViewModel();
            ReloadDefaultsFromConfig();
            ReloadAvailableSchedules();
            ReloadMappingsForSelectedSchedule();
        }

        public Task<ExcelScheduleSyncResult> SyncAsync(
            ExcelSyncFromAppRequest request,
            CancellationToken cancellationToken)
        {
            return _app.SyncExcelAsync(request, cancellationToken);
        }

        public void ReloadAvailableSchedules()
        {
            ViewModel.SetScheduleNames(_app.GetAvailableSchedulesForExcelSync());
        }

        public void ReloadDefaultsFromConfig()
        {
            var config = _app.LoadConfig();
            ViewModel.SetDefaults(
                config.Excel.DefaultDirectory,
                config.Excel.DefaultWorksheetName,
                config.Excel.AnchorColumn);
            _protectedColumns = config.Excel.ProtectedSystemColumns.ToArray();
            ViewModel.ApplyProtectedColumns(_protectedColumns);
        }

        public bool? ShowDialog()
        {
            Window window = CreateWindow();

            ComboBox directionCombo = new ComboBox
            {
                Width = 180,
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

            TextBox fileTextBox = new TextBox
            {
                Margin = new Thickness(8, 0, 8, 0),
                MinWidth = 560,
                Text = ViewModel.FilePath,
            };

            Button browseButton = new Button
            {
                Width = 80,
                Content = "Browse",
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
                IsChecked = ViewModel.PreviewOnly,
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
                Content = "Sync Excel",
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
                ReloadMappingsForSelectedSchedule();
                RefreshMappingsGrid(mappingsGrid);
                worksheetTextBox.Text = ViewModel.WorksheetName;
                fileTextBox.Text = ViewModel.FilePath;
            };

            scheduleCombo.SelectionChanged += (_, _) =>
            {
                ViewModel.SelectedScheduleName = scheduleCombo.SelectedItem as string ?? string.Empty;
                ViewModel.ApplySelectedScheduleDefaults();
                ReloadMappingsForSelectedSchedule();
                RefreshMappingsGrid(mappingsGrid);
                worksheetTextBox.Text = ViewModel.WorksheetName;
                fileTextBox.Text = ViewModel.FilePath;
            };

            browseButton.Click += (_, _) =>
            {
                ApplyControlValues(
                    directionCombo,
                    scheduleCombo,
                    fileTextBox,
                    worksheetTextBox,
                    anchorTextBox,
                    previewCheckBox);

                string selected = BrowseForWorkbook(ViewModel.Direction, ViewModel.FilePath);
                if (!string.IsNullOrWhiteSpace(selected))
                {
                    ViewModel.FilePath = selected;
                    fileTextBox.Text = selected;
                }
            };

            syncButton.Click += (_, _) =>
            {
                _ = RunSyncAsync(
                    window,
                    directionCombo,
                    scheduleCombo,
                    fileTextBox,
                    worksheetTextBox,
                    anchorTextBox,
                    previewCheckBox,
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
                fileTextBox,
                browseButton,
                worksheetTextBox,
                anchorTextBox,
                previewCheckBox,
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
                Title = "Excel Sync",
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

            grid.Columns.Add(new DataGridTextColumn
            {
                Header = "Excel Column",
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                Binding = new Binding(nameof(ExcelSyncMappingItem.ExcelColumn)),
            });
            grid.Columns.Add(new DataGridTextColumn
            {
                Header = "Revit Parameter",
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                Binding = new Binding(nameof(ExcelSyncMappingItem.RevitParameter)),
            });
            grid.Columns.Add(new DataGridCheckBoxColumn
            {
                Header = "Editable",
                Width = 120,
                Binding = new Binding(nameof(ExcelSyncMappingItem.IsEditable)),
            });

            return grid;
        }

        private Grid BuildLayout(
            ComboBox directionCombo,
            ComboBox scheduleCombo,
            TextBox fileTextBox,
            Button browseButton,
            TextBox worksheetTextBox,
            TextBox anchorTextBox,
            CheckBox previewCheckBox,
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
                Text = "Sync Revit Schedule <-> Excel",
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

            StackPanel fileRow = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 0, 0, 8),
            };
            fileRow.Children.Add(new TextBlock
            {
                Width = 110,
                VerticalAlignment = VerticalAlignment.Center,
                Text = "Excel File",
            });
            fileRow.Children.Add(fileTextBox);
            fileRow.Children.Add(browseButton);
            formGrid.Children.Add(fileRow);
            Grid.SetRow(fileRow, 1);
            Grid.SetColumn(fileRow, 0);
            Grid.SetColumnSpan(fileRow, 2);

            formGrid.Children.Add(CreateLabeledRow("Worksheet", worksheetTextBox, 2, 0, 8));
            formGrid.Children.Add(CreateLabeledRow("Anchor Column", anchorTextBox, 2, 1, 0));

            StackPanel options = new StackPanel
            {
                Orientation = Orientation.Horizontal,
            };
            options.Children.Add(previewCheckBox);
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
            double rightMargin)
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
            return panel;
        }

        private void ReloadMappingsForSelectedSchedule()
        {
            IReadOnlyList<GoogleSheetColumnMapping> mappings =
                _app.GetScheduleColumnMappingsForExcelSync(ViewModel.SelectedScheduleName);
            ViewModel.SetColumnMappings(mappings);
            ViewModel.ApplyProtectedColumns(_protectedColumns);
        }

        private static string BrowseForWorkbook(string direction, string currentPath)
        {
            bool import = string.Equals(direction, GoogleSyncDirections.Import, StringComparison.OrdinalIgnoreCase);
            if (import)
            {
                OpenFileDialog open = new OpenFileDialog
                {
                    Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                    FileName = currentPath ?? string.Empty,
                    CheckFileExists = true,
                };
                return open.ShowDialog() == true ? open.FileName : string.Empty;
            }

            SaveFileDialog save = new SaveFileDialog
            {
                Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                FileName = currentPath ?? string.Empty,
                AddExtension = true,
                DefaultExt = ".xlsx",
                OverwritePrompt = true,
            };
            return save.ShowDialog() == true ? save.FileName : string.Empty;
        }

        private void RefreshMappingsGrid(DataGrid mappingsGrid)
        {
            mappingsGrid.ItemsSource = null;
            mappingsGrid.ItemsSource = ViewModel.ColumnMappings;
        }

        private void ApplyControlValues(
            ComboBox directionCombo,
            ComboBox scheduleCombo,
            TextBox fileTextBox,
            TextBox worksheetTextBox,
            TextBox anchorTextBox,
            CheckBox previewCheckBox)
        {
            ViewModel.Direction = directionCombo.SelectedItem as string ?? GoogleSyncDirections.Export;
            ViewModel.SelectedScheduleName = scheduleCombo.SelectedItem as string ?? string.Empty;
            ViewModel.FilePath = fileTextBox.Text ?? string.Empty;
            ViewModel.WorksheetName = worksheetTextBox.Text ?? string.Empty;
            ViewModel.AnchorColumn = anchorTextBox.Text ?? "MDR_UNIQUE_ID";
            ViewModel.PreviewOnly = previewCheckBox.IsChecked ?? false;
        }

        private async Task RunSyncAsync(
            Window window,
            ComboBox directionCombo,
            ComboBox scheduleCombo,
            TextBox fileTextBox,
            TextBox worksheetTextBox,
            TextBox anchorTextBox,
            CheckBox previewCheckBox,
            Button refreshButton,
            Button syncButton,
            Button closeButton)
        {
            string originalSyncText = syncButton.Content?.ToString() ?? "Sync Excel";
            try
            {
                syncButton.IsEnabled = false;
                refreshButton.IsEnabled = false;
                closeButton.IsEnabled = false;
                syncButton.Content = "Syncing...";

                ApplyControlValues(
                    directionCombo,
                    scheduleCombo,
                    fileTextBox,
                    worksheetTextBox,
                    anchorTextBox,
                    previewCheckBox);

                if (string.Equals(ViewModel.Direction, GoogleSyncDirections.Import, StringComparison.OrdinalIgnoreCase))
                {
                    ExcelScheduleSyncResult previewResult = await RunImportAsync(previewOnly: true, password: string.Empty)
                        .ConfigureAwait(true);
                    DiffViewerWindow previewWindow = new DiffViewerWindow();
                    previewWindow.SetDiff(previewResult.DiffResult);

                    bool shouldApply = previewWindow.ShowDialog(allowApply: !ViewModel.PreviewOnly);
                    if (ViewModel.PreviewOnly || !shouldApply)
                    {
                        ShowSyncSummary(previewResult);
                        return;
                    }

                    string password = PromptForPassword();
                    if (string.IsNullOrWhiteSpace(password))
                    {
                        return;
                    }

                    ExcelScheduleSyncResult applyResult = await RunImportAsync(previewOnly: false, password: password)
                        .ConfigureAwait(true);
                    ShowSyncSummary(applyResult);
                    window.DialogResult = true;
                    window.Close();
                    return;
                }

                ExcelScheduleSyncResult exportResult = await SyncAsync(ViewModel.BuildRequest(), CancellationToken.None)
                    .ConfigureAwait(true);
                ShowSyncSummary(exportResult);
                window.DialogResult = true;
                window.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message,
                    "Excel Sync Failed",
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

        private async Task<ExcelScheduleSyncResult> RunImportAsync(bool previewOnly, string password)
        {
            ExcelSyncFromAppRequest request = ViewModel.BuildRequest();
            request.Direction = GoogleSyncDirections.Import;
            request.PreviewOnly = previewOnly;
            request.ImportPassword = password ?? string.Empty;
            return await SyncAsync(request, CancellationToken.None).ConfigureAwait(true);
        }

        private static string PromptForPassword()
        {
            Window window = new Window
            {
                Title = "Excel Import Password",
                Width = 360,
                Height = 170,
                ResizeMode = ResizeMode.NoResize,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
            };

            Grid root = new Grid
            {
                Margin = new Thickness(16),
            };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            TextBlock label = new TextBlock
            {
                Margin = new Thickness(0, 0, 0, 8),
                Text = "Enter Excel import password.",
            };
            root.Children.Add(label);
            Grid.SetRow(label, 0);

            PasswordBox passwordBox = new PasswordBox
            {
                Margin = new Thickness(0, 0, 0, 12),
            };
            root.Children.Add(passwordBox);
            Grid.SetRow(passwordBox, 1);

            StackPanel buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
            };
            Button ok = new Button
            {
                Width = 80,
                Margin = new Thickness(0, 0, 8, 0),
                Content = "Apply",
                IsDefault = true,
            };
            Button cancel = new Button
            {
                Width = 80,
                Content = "Cancel",
                IsCancel = true,
            };
            buttons.Children.Add(ok);
            buttons.Children.Add(cancel);
            root.Children.Add(buttons);
            Grid.SetRow(buttons, 2);

            ok.Click += (_, _) =>
            {
                window.DialogResult = true;
                window.Close();
            };
            cancel.Click += (_, _) =>
            {
                window.DialogResult = false;
                window.Close();
            };

            window.Content = root;
            return window.ShowDialog() == true ? passwordBox.Password ?? string.Empty : string.Empty;
        }

        private static void ShowSyncSummary(ExcelScheduleSyncResult result)
        {
            string message;
            if (string.Equals(result.Direction, GoogleSyncDirections.Export, StringComparison.OrdinalIgnoreCase))
            {
                List<string> lines = new List<string>
                {
                    "Export completed.",
                    "Rows exported: " + result.ExportedRows,
                    "Rows skipped: " + result.SkippedRows,
                };

                if (!string.IsNullOrWhiteSpace(result.WriteResult.FilePath))
                {
                    lines.Add("File: " + result.WriteResult.FilePath);
                }

                message = string.Join("\n", lines);
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
                "Excel Sync",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
    }
}
