using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Threading;
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
            ReloadDefaultsFromConfig();
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

        public bool? ShowDialog()
        {
            Window window = CreateWindow();

            TextBox usernameTextBox = new TextBox
            {
                Width = 240,
                Margin = new Thickness(8, 0, 16, 0),
                Text = ViewModel.Username,
            };
            PasswordBox passwordBox = new PasswordBox
            {
                Width = 240,
                Margin = new Thickness(8, 0, 0, 0),
            };
            passwordBox.Password = ViewModel.Password;

            TextBox projectCodeTextBox = new TextBox
            {
                Width = 240,
                Margin = new Thickness(8, 0, 16, 0),
                Text = ViewModel.ProjectCode,
            };
            TextBox outputDirectoryTextBox = new TextBox
            {
                Margin = new Thickness(8, 0, 0, 0),
                MinWidth = 420,
                Text = ViewModel.OutputDirectory,
            };
            TextBox apiBaseUrlTextBox = new TextBox
            {
                Margin = new Thickness(8, 0, 8, 0),
                MinWidth = 420,
                IsReadOnly = true,
                Text = ViewModel.ApiBaseUrl,
            };
            Button openConfigButton = new Button
            {
                Width = 100,
                Content = "Open Config",
            };

            CheckBox includeNativeCheckBox = new CheckBox
            {
                Content = "Include native export",
                Margin = new Thickness(0, 0, 16, 0),
                IsChecked = ViewModel.IncludeNative,
            };
            CheckBox retryFailedCheckBox = new CheckBox
            {
                Content = "Retry failed items",
                IsChecked = ViewModel.RetryFailedItems,
            };

            TextBlock selectedCountText = new TextBlock
            {
                Margin = new Thickness(8, 0, 0, 0),
                VerticalAlignment = VerticalAlignment.Center,
                FontWeight = FontWeights.SemiBold,
                Text = ViewModel.SelectedCount.ToString(CultureInfo.InvariantCulture),
            };

            DataGrid sheetsGrid = BuildSheetsGrid();
            sheetsGrid.ItemsSource = ViewModel.SheetRows;

            Button refreshButton = new Button
            {
                Width = 120,
                Margin = new Thickness(0, 0, 8, 0),
                Content = "Refresh Sheets",
            };
            Button publishButton = new Button
            {
                Width = 100,
                Margin = new Thickness(0, 0, 8, 0),
                IsDefault = true,
                Content = "Publish",
            };
            Button cancelButton = new Button
            {
                Width = 90,
                IsCancel = true,
                Content = "Cancel",
            };

            Grid root = BuildLayout(
                usernameTextBox,
                passwordBox,
                apiBaseUrlTextBox,
                openConfigButton,
                projectCodeTextBox,
                outputDirectoryTextBox,
                includeNativeCheckBox,
                retryFailedCheckBox,
                selectedCountText,
                sheetsGrid,
                refreshButton,
                publishButton,
                cancelButton);
            window.Content = root;
            bool hasPublished = false;

            void RefreshSelectedCountText()
            {
                selectedCountText.Text = ViewModel.SelectedCount.ToString(CultureInfo.InvariantCulture);
            }

            void RefreshSheetBindings()
            {
                sheetsGrid.ItemsSource = null;
                sheetsGrid.ItemsSource = ViewModel.SheetRows;
                RefreshSelectedCountText();
            }

            void ApplyControlValues()
            {
                ViewModel.Username = usernameTextBox.Text ?? string.Empty;
                ViewModel.Password = passwordBox.Password ?? string.Empty;
                ViewModel.ProjectCode = projectCodeTextBox.Text ?? string.Empty;
                ViewModel.OutputDirectory = outputDirectoryTextBox.Text ?? string.Empty;
                ViewModel.IncludeNative = includeNativeCheckBox.IsChecked ?? false;
                ViewModel.RetryFailedItems = retryFailedCheckBox.IsChecked ?? true;
                ViewModel.ApiBaseUrl = apiBaseUrlTextBox.Text ?? string.Empty;
            }

            sheetsGrid.CellEditEnding += (_, _) =>
            {
                window.Dispatcher.BeginInvoke(
                    DispatcherPriority.Background,
                    new Action(RefreshSelectedCountText));
            };
            sheetsGrid.CurrentCellChanged += (_, _) => RefreshSelectedCountText();

            refreshButton.Click += (_, _) =>
            {
                ReloadAvailableSheets();
                RefreshSheetBindings();
            };

            openConfigButton.Click += (_, _) =>
            {
                OpenConfig();
                ReloadDefaultsFromConfig();
                projectCodeTextBox.Text = ViewModel.ProjectCode;
                outputDirectoryTextBox.Text = ViewModel.OutputDirectory;
                includeNativeCheckBox.IsChecked = ViewModel.IncludeNative;
                retryFailedCheckBox.IsChecked = ViewModel.RetryFailedItems;
                apiBaseUrlTextBox.Text = ViewModel.ApiBaseUrl;
            };

            publishButton.Click += (_, _) =>
            {
                try
                {
                    ApplyControlValues();

                    if (string.IsNullOrWhiteSpace(ViewModel.Username) || string.IsNullOrWhiteSpace(ViewModel.Password))
                    {
                        MessageBox.Show(
                            "Username and password are required.",
                            "Publish to MDR",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);
                        return;
                    }

                    if (string.IsNullOrWhiteSpace(ViewModel.ProjectCode))
                    {
                        MessageBox.Show(
                            "Project code is required.",
                            "Publish to MDR",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);
                        return;
                    }

                    IReadOnlyList<PublishSheetItem> selectedItems = ViewModel.BuildSelectedItems();
                    if (selectedItems.Count == 0)
                    {
                        MessageBox.Show(
                            "No sheets are selected. Select at least one sheet to publish.",
                            "Publish to MDR",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);
                        return;
                    }

                    PublishFromAppRequest request = new PublishFromAppRequest
                    {
                        Username = ViewModel.Username,
                        Password = ViewModel.Password,
                        ProjectCode = ViewModel.ProjectCode,
                        OutputDirectory = ViewModel.OutputDirectory,
                        IncludeNative = ViewModel.IncludeNative,
                        RetryFailedItems = ViewModel.RetryFailedItems,
                    };

                    for (int i = 0; i < selectedItems.Count; i++)
                    {
                        request.Items.Add(selectedItems[i]);
                    }

                    PublishSheetsCommandResult result = PublishAsync(request, CancellationToken.None)
                        .GetAwaiter()
                        .GetResult();
                    ViewModel.ApplyPublishResult(result.FinalResponse);
                    RefreshSheetBindings();
                    hasPublished = true;

                    MessageBox.Show(
                        BuildSummaryMessage(result),
                        "Publish to MDR",
                        MessageBoxButton.OK,
                        result.FinalResponse.Summary.FailedCount > 0
                            ? MessageBoxImage.Warning
                            : MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        ex.Message,
                        "Publish to MDR Failed",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                }
            };

            cancelButton.Click += (_, _) =>
            {
                window.DialogResult = hasPublished ? true : false;
                window.Close();
            };

            return window.ShowDialog();
        }

        private void ReloadDefaultsFromConfig()
        {
            var config = _app.LoadConfig();
            ViewModel.ApiBaseUrl = config.ApiBaseUrl ?? string.Empty;
            ViewModel.ProjectCode = config.ProjectCode ?? string.Empty;
            ViewModel.OutputDirectory = config.PublishOutputDirectory ?? string.Empty;
            ViewModel.IncludeNative = config.IncludeNativeByDefault;
            ViewModel.RetryFailedItems = config.RetryFailedItems;
        }

        private void OpenConfig()
        {
            try
            {
                ProcessStartInfo processInfo = new ProcessStartInfo
                {
                    FileName = "notepad.exe",
                    Arguments = "\"" + _app.ConfigPath + "\"",
                    UseShellExecute = true,
                };
                Process.Start(processInfo);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message,
                    "Open Config Failed",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private static Window CreateWindow()
        {
            return new Window
            {
                Title = "Publish to MDR",
                Width = 980,
                Height = 640,
                MinWidth = 760,
                MinHeight = 520,
                ResizeMode = ResizeMode.CanResize,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
            };
        }

        private static DataGrid BuildSheetsGrid()
        {
            DataGrid grid = new DataGrid
            {
                AutoGenerateColumns = false,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                SelectionMode = DataGridSelectionMode.Extended,
                SelectionUnit = DataGridSelectionUnit.FullRow,
                Margin = new Thickness(0, 0, 0, 12),
            };

            grid.Columns.Add(new DataGridCheckBoxColumn
            {
                Header = "Publish",
                Width = 90,
                Binding = new Binding(nameof(PublishSheetSelectionItem.IsSelected))
                {
                    Mode = BindingMode.TwoWay,
                    UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                },
            });
            grid.Columns.Add(new DataGridTextColumn
            {
                Header = "Sheet Number",
                Width = 170,
                Binding = new Binding(nameof(PublishSheetSelectionItem.SheetNumber)),
            });
            grid.Columns.Add(new DataGridTextColumn
            {
                Header = "Sheet Name",
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                Binding = new Binding(nameof(PublishSheetSelectionItem.SheetName)),
            });
            grid.Columns.Add(new DataGridTextColumn
            {
                Header = "Revision",
                Width = 120,
                Binding = new Binding(nameof(PublishSheetSelectionItem.RequestedRevision)),
            });
            grid.Columns.Add(new DataGridTextColumn
            {
                Header = "Doc Status",
                Width = 100,
                Binding = new Binding(nameof(PublishSheetSelectionItem.StatusCode)),
            });
            grid.Columns.Add(new DataGridTextColumn
            {
                Header = "Status",
                Width = 120,
                IsReadOnly = true,
                Binding = new Binding(nameof(PublishSheetSelectionItem.LastState)),
            });
            grid.Columns.Add(new DataGridTextColumn
            {
                Header = "Error",
                Width = 160,
                IsReadOnly = true,
                Binding = new Binding(nameof(PublishSheetSelectionItem.LastErrorCode)),
            });

            return grid;
        }

        private static Grid BuildLayout(
            TextBox usernameTextBox,
            PasswordBox passwordBox,
            TextBox apiBaseUrlTextBox,
            Button openConfigButton,
            TextBox projectCodeTextBox,
            TextBox outputDirectoryTextBox,
            CheckBox includeNativeCheckBox,
            CheckBox retryFailedCheckBox,
            TextBlock selectedCountText,
            DataGrid sheetsGrid,
            Button refreshButton,
            Button publishButton,
            Button cancelButton)
        {
            Grid root = new Grid
            {
                Margin = new Thickness(16),
            };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            TextBlock info = new TextBlock
            {
                Margin = new Thickness(0, 0, 0, 8),
                FontWeight = FontWeights.SemiBold,
                Text = "Select sheets and publish PDF/native files to MDR EDMS",
            };
            root.Children.Add(info);
            Grid.SetRow(info, 0);

            Grid authRow = new Grid
            {
                Margin = new Thickness(0, 0, 0, 8),
            };
            authRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            authRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            StackPanel userPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
            };
            userPanel.Children.Add(new TextBlock
            {
                Width = 95,
                VerticalAlignment = VerticalAlignment.Center,
                Text = "Username",
            });
            userPanel.Children.Add(usernameTextBox);
            authRow.Children.Add(userPanel);
            Grid.SetColumn(userPanel, 0);

            StackPanel passPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
            };
            passPanel.Children.Add(new TextBlock
            {
                Width = 95,
                VerticalAlignment = VerticalAlignment.Center,
                Text = "Password",
            });
            passPanel.Children.Add(passwordBox);
            authRow.Children.Add(passPanel);
            Grid.SetColumn(passPanel, 1);

            root.Children.Add(authRow);
            Grid.SetRow(authRow, 1);

            Grid optionsRow = new Grid
            {
                Margin = new Thickness(0, 0, 0, 8),
            };
            optionsRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            optionsRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            StackPanel projectPanel = new StackPanel
            {
                Orientation = Orientation.Vertical,
            };

            StackPanel apiPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 0, 0, 6),
            };
            apiPanel.Children.Add(new TextBlock
            {
                Width = 95,
                VerticalAlignment = VerticalAlignment.Center,
                Text = "API Base URL",
            });
            apiPanel.Children.Add(apiBaseUrlTextBox);
            apiPanel.Children.Add(openConfigButton);
            projectPanel.Children.Add(apiPanel);

            StackPanel projectCodePanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 0, 0, 6),
            };
            projectCodePanel.Children.Add(new TextBlock
            {
                Width = 95,
                VerticalAlignment = VerticalAlignment.Center,
                Text = "Project Code",
            });
            projectCodePanel.Children.Add(projectCodeTextBox);
            projectPanel.Children.Add(projectCodePanel);

            StackPanel outputPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
            };
            outputPanel.Children.Add(new TextBlock
            {
                Width = 95,
                VerticalAlignment = VerticalAlignment.Center,
                Text = "Output Dir",
            });
            outputPanel.Children.Add(outputDirectoryTextBox);
            projectPanel.Children.Add(outputPanel);

            optionsRow.Children.Add(projectPanel);
            Grid.SetColumn(projectPanel, 0);

            StackPanel switchesPanel = new StackPanel
            {
                Orientation = Orientation.Vertical,
                Margin = new Thickness(16, 0, 0, 0),
            };
            switchesPanel.Children.Add(includeNativeCheckBox);
            switchesPanel.Children.Add(retryFailedCheckBox);

            StackPanel selectedPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 8, 0, 0),
            };
            selectedPanel.Children.Add(new TextBlock
            {
                VerticalAlignment = VerticalAlignment.Center,
                Text = "Selected Sheets:",
            });
            selectedPanel.Children.Add(selectedCountText);
            switchesPanel.Children.Add(selectedPanel);

            optionsRow.Children.Add(switchesPanel);
            Grid.SetColumn(switchesPanel, 1);

            root.Children.Add(optionsRow);
            Grid.SetRow(optionsRow, 2);

            root.Children.Add(sheetsGrid);
            Grid.SetRow(sheetsGrid, 3);

            StackPanel buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
            };
            buttons.Children.Add(refreshButton);
            buttons.Children.Add(publishButton);
            buttons.Children.Add(cancelButton);
            root.Children.Add(buttons);
            Grid.SetRow(buttons, 4);

            return root;
        }

        private static string BuildSummaryMessage(PublishSheetsCommandResult result)
        {
            PublishBatchResponse response = result.FinalResponse;

            StringBuilder builder = new StringBuilder();
            builder.AppendLine("Publish completed.");
            builder.AppendLine("Run ID: " + (response.RunId ?? string.Empty));
            builder.AppendLine("Status: " + (response.Summary.Status ?? string.Empty));
            builder.AppendLine("Requested: " + response.Summary.RequestedCount.ToString(CultureInfo.InvariantCulture));
            builder.AppendLine("Success: " + response.Summary.SuccessCount.ToString(CultureInfo.InvariantCulture));
            builder.AppendLine("Failed: " + response.Summary.FailedCount.ToString(CultureInfo.InvariantCulture));
            builder.AppendLine("Duplicate: " + response.Summary.DuplicateCount.ToString(CultureInfo.InvariantCulture));

            if (!string.IsNullOrWhiteSpace(result.OutputDirectory))
            {
                builder.AppendLine("Output: " + result.OutputDirectory);
            }

            string failures = BuildFailedItemsSection(response);
            if (!string.IsNullOrWhiteSpace(failures))
            {
                builder.AppendLine();
                builder.Append(failures);
            }

            return builder.ToString();
        }

        private static string BuildFailedItemsSection(PublishBatchResponse response)
        {
            if (response.Items == null || response.Items.Count == 0)
            {
                return string.Empty;
            }

            StringBuilder builder = new StringBuilder();
            int shown = 0;
            for (int i = 0; i < response.Items.Count; i++)
            {
                PublishItemResult item = response.Items[i];
                if (item == null || !item.IsFailed())
                {
                    continue;
                }

                if (shown == 0)
                {
                    builder.AppendLine("Failed items:");
                }

                builder.Append(" - #");
                builder.Append(item.ItemIndex.ToString(CultureInfo.InvariantCulture));
                builder.Append(" ");
                builder.Append(string.IsNullOrWhiteSpace(item.ErrorCode) ? "failed" : item.ErrorCode);
                if (!string.IsNullOrWhiteSpace(item.ErrorMessage))
                {
                    builder.Append(": ");
                    builder.Append(item.ErrorMessage);
                }
                builder.AppendLine();

                shown++;
                if (shown >= 10)
                {
                    builder.AppendLine(" - ...");
                    break;
                }
            }

            return builder.ToString().TrimEnd();
        }
    }
}
