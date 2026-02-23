using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Threading;
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
            ViewModel = new SmartNumberingWindowViewModel();
            ReloadRules();
        }

        public SmartNumberingWindowViewModel ViewModel { get; }

        public SmartNumberingResult Apply(SmartNumberingFromAppRequest request)
        {
            return _app.ApplySmartNumbering(request);
        }

        public bool? ShowDialog()
        {
            Window window = new Window
            {
                Title = "Smart Numbering",
                Width = 980,
                Height = 650,
                MinWidth = 820,
                MinHeight = 520,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                ResizeMode = ResizeMode.CanResize,
            };

            ComboBox ruleCombo = new ComboBox
            {
                Width = 260,
                Margin = new Thickness(8, 0, 8, 0),
                DisplayMemberPath = nameof(SmartNumberingRuleOption.DisplayName),
                SelectedValuePath = nameof(SmartNumberingRuleOption.RuleId),
                ItemsSource = ViewModel.RuleOptions,
                SelectedValue = ViewModel.SelectedRuleId,
            };

            TextBox formulaTextBox = new TextBox
            {
                Margin = new Thickness(8, 0, 0, 0),
                MinWidth = 640,
                Text = ViewModel.Formula,
            };

            TextBox targetsTextBox = new TextBox
            {
                Margin = new Thickness(8, 0, 0, 0),
                MinWidth = 640,
                Text = ViewModel.TargetsText,
            };

            TextBox startAtTextBox = new TextBox
            {
                Width = 80,
                Margin = new Thickness(8, 0, 24, 0),
                Text = ViewModel.StartAt.ToString(CultureInfo.InvariantCulture),
            };

            TextBox widthTextBox = new TextBox
            {
                Width = 80,
                Margin = new Thickness(8, 0, 0, 0),
                Text = ViewModel.SequenceWidth.ToString(CultureInfo.InvariantCulture),
            };

            TextBlock statusText = new TextBlock
            {
                Margin = new Thickness(0, 0, 0, 10),
                Text = ViewModel.StatusText,
            };

            DataGrid previewGrid = BuildPreviewGrid();
            previewGrid.ItemsSource = ViewModel.PreviewRows;

            Button previewButton = new Button
            {
                Width = 100,
                Margin = new Thickness(0, 0, 8, 0),
                Content = "Preview",
                IsDefault = true,
            };

            Button applyButton = new Button
            {
                Width = 100,
                Margin = new Thickness(0, 0, 8, 0),
                Content = "Apply",
                IsEnabled = ViewModel.CanApply,
            };

            Button closeButton = new Button
            {
                Width = 90,
                Content = "Close",
                IsCancel = true,
            };

            Grid root = BuildLayout(
                ruleCombo,
                formulaTextBox,
                targetsTextBox,
                startAtTextBox,
                widthTextBox,
                statusText,
                previewGrid,
                previewButton,
                applyButton,
                closeButton);
            window.Content = root;

            bool previewDirty = true;
            DispatcherTimer debounceTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(350),
            };

            void RefreshPreviewBindings()
            {
                previewGrid.ItemsSource = null;
                previewGrid.ItemsSource = ViewModel.PreviewRows;
                statusText.Text = ViewModel.StatusText;
                applyButton.IsEnabled = ViewModel.CanApply;
            }

            bool TryApplyControlValues(bool fromRuleSelection, out string validationMessage)
            {
                validationMessage = string.Empty;

                ViewModel.SelectedRuleId = ruleCombo.SelectedValue as string ?? string.Empty;
                if (fromRuleSelection)
                {
                    ViewModel.ApplySelectedRuleTemplate();
                    formulaTextBox.Text = ViewModel.Formula;
                    targetsTextBox.Text = ViewModel.TargetsText;
                    startAtTextBox.Text = ViewModel.StartAt.ToString(CultureInfo.InvariantCulture);
                    widthTextBox.Text = ViewModel.SequenceWidth.ToString(CultureInfo.InvariantCulture);
                }
                else
                {
                    ViewModel.Formula = formulaTextBox.Text ?? string.Empty;
                    ViewModel.TargetsText = targetsTextBox.Text ?? string.Empty;
                    if (!int.TryParse(startAtTextBox.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int startAt) || startAt <= 0)
                    {
                        validationMessage = "StartAt must be a positive integer.";
                        return false;
                    }

                    if (!int.TryParse(widthTextBox.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int width) || width <= 0)
                    {
                        validationMessage = "SequenceWidth must be a positive integer.";
                        return false;
                    }

                    ViewModel.StartAt = startAt;
                    ViewModel.SequenceWidth = width;
                }

                return true;
            }

            SmartNumberingResult Run(bool previewOnly)
            {
                SmartNumberingFromAppRequest request = new SmartNumberingFromAppRequest
                {
                    Rule = ViewModel.BuildRule(),
                    PreviewOnly = previewOnly,
                };
                return Apply(request);
            }

            void RunPreview()
            {
                if (!TryApplyControlValues(fromRuleSelection: false, out string message))
                {
                    ViewModel.UpdatePreview(new SmartNumberingResult
                    {
                        FatalErrorCode = "validation_error",
                        FatalErrorMessage = message,
                    });
                    statusText.Text = message;
                    applyButton.IsEnabled = false;
                    return;
                }

                SmartNumberingResult result = Run(previewOnly: true);
                ViewModel.UpdatePreview(result);
                RefreshPreviewBindings();
                previewDirty = false;
            }

            void MarkDirty()
            {
                previewDirty = true;
                debounceTimer.Stop();
                debounceTimer.Start();
            }

            debounceTimer.Tick += (_, _) =>
            {
                debounceTimer.Stop();
                if (!previewDirty)
                {
                    return;
                }

                RunPreview();
            };

            ruleCombo.SelectionChanged += (_, _) =>
            {
                if (!TryApplyControlValues(fromRuleSelection: true, out _))
                {
                    return;
                }

                MarkDirty();
            };
            formulaTextBox.TextChanged += (_, _) => MarkDirty();
            targetsTextBox.TextChanged += (_, _) => MarkDirty();
            startAtTextBox.TextChanged += (_, _) => MarkDirty();
            widthTextBox.TextChanged += (_, _) => MarkDirty();

            previewButton.Click += (_, _) => RunPreview();
            applyButton.Click += (_, _) =>
            {
                try
                {
                    if (previewDirty)
                    {
                        RunPreview();
                    }

                    if (!ViewModel.CanApply)
                    {
                        MessageBox.Show(
                            "Fix preview errors before applying changes.",
                            "Smart Numbering",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);
                        return;
                    }

                    SmartNumberingResult result = Run(previewOnly: false);
                    ViewModel.UpdatePreview(result);
                    RefreshPreviewBindings();

                    string message = "Applied: " + result.AppliedCount + "\n" +
                        "Skipped: " + result.SkippedCount + "\n" +
                        "Failed: " + result.FailedCount;
                    if (!string.IsNullOrWhiteSpace(result.FatalErrorCode))
                    {
                        message += "\nStatus: " + result.FatalErrorCode;
                    }

                    MessageBox.Show(
                        message,
                        "Smart Numbering",
                        MessageBoxButton.OK,
                        result.AppliedCount > 0 ? MessageBoxImage.Information : MessageBoxImage.Warning);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        ex.Message,
                        "Smart Numbering Failed",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                }
            };

            closeButton.Click += (_, _) =>
            {
                window.DialogResult = false;
                window.Close();
            };

            // Initial preview is run once with debounce policy active.
            MarkDirty();
            return window.ShowDialog();
        }

        private void ReloadRules()
        {
            var config = _app.LoadConfig();
            ViewModel.SetRules(config.SmartNumbering.Rules, config.SmartNumbering.DefaultRuleId);
        }

        private static DataGrid BuildPreviewGrid()
        {
            DataGrid grid = new DataGrid
            {
                AutoGenerateColumns = false,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                Margin = new Thickness(0, 0, 0, 10),
            };

            grid.Columns.Add(new DataGridTextColumn
            {
                Header = "Element",
                Width = 220,
                Binding = new Binding(nameof(SmartNumberingPreviewRow.Element)),
            });
            grid.Columns.Add(new DataGridTextColumn
            {
                Header = "Target",
                Width = 180,
                Binding = new Binding(nameof(SmartNumberingPreviewRow.Target)),
            });
            grid.Columns.Add(new DataGridTextColumn
            {
                Header = "Current",
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                Binding = new Binding(nameof(SmartNumberingPreviewRow.Current)),
            });
            grid.Columns.Add(new DataGridTextColumn
            {
                Header = "Proposed",
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                Binding = new Binding(nameof(SmartNumberingPreviewRow.Proposed)),
            });
            grid.Columns.Add(new DataGridTextColumn
            {
                Header = "Status",
                Width = 110,
                Binding = new Binding(nameof(SmartNumberingPreviewRow.Status)),
            });
            grid.Columns.Add(new DataGridTextColumn
            {
                Header = "Error",
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                Binding = new Binding(nameof(SmartNumberingPreviewRow.Error)),
            });

            return grid;
        }

        private static Grid BuildLayout(
            ComboBox ruleCombo,
            TextBox formulaTextBox,
            TextBox targetsTextBox,
            TextBox startAtTextBox,
            TextBox widthTextBox,
            TextBlock statusText,
            DataGrid previewGrid,
            Button previewButton,
            Button applyButton,
            Button closeButton)
        {
            Grid root = new Grid
            {
                Margin = new Thickness(16),
            };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            TextBlock title = new TextBlock
            {
                Margin = new Thickness(0, 0, 0, 8),
                FontWeight = FontWeights.SemiBold,
                Text = "Rule-based element numbering with atomic apply",
            };
            root.Children.Add(title);
            Grid.SetRow(title, 0);

            StackPanel rulePanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 0, 0, 8),
            };
            rulePanel.Children.Add(new TextBlock
            {
                Width = 120,
                VerticalAlignment = VerticalAlignment.Center,
                Text = "Rule",
            });
            rulePanel.Children.Add(ruleCombo);
            root.Children.Add(rulePanel);
            Grid.SetRow(rulePanel, 1);

            StackPanel formulaPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 0, 0, 8),
            };
            formulaPanel.Children.Add(new TextBlock
            {
                Width = 120,
                VerticalAlignment = VerticalAlignment.Center,
                Text = "Formula",
            });
            formulaPanel.Children.Add(formulaTextBox);
            root.Children.Add(formulaPanel);
            Grid.SetRow(formulaPanel, 2);

            Grid optionsRow = new Grid
            {
                Margin = new Thickness(0, 0, 0, 8),
            };
            optionsRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            optionsRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            optionsRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            StackPanel targetsPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
            };
            targetsPanel.Children.Add(new TextBlock
            {
                Width = 120,
                VerticalAlignment = VerticalAlignment.Center,
                Text = "Targets",
            });
            targetsPanel.Children.Add(targetsTextBox);
            optionsRow.Children.Add(targetsPanel);
            Grid.SetColumn(targetsPanel, 0);

            StackPanel startPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
            };
            startPanel.Children.Add(new TextBlock
            {
                Width = 60,
                VerticalAlignment = VerticalAlignment.Center,
                Text = "StartAt",
            });
            startPanel.Children.Add(startAtTextBox);
            optionsRow.Children.Add(startPanel);
            Grid.SetColumn(startPanel, 1);

            StackPanel widthPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
            };
            widthPanel.Children.Add(new TextBlock
            {
                Width = 95,
                VerticalAlignment = VerticalAlignment.Center,
                Text = "SequenceWidth",
            });
            widthPanel.Children.Add(widthTextBox);
            optionsRow.Children.Add(widthPanel);
            Grid.SetColumn(widthPanel, 2);

            root.Children.Add(optionsRow);
            Grid.SetRow(optionsRow, 3);

            root.Children.Add(previewGrid);
            Grid.SetRow(previewGrid, 4);

            Grid footer = new Grid();
            footer.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            footer.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            footer.Children.Add(statusText);
            Grid.SetColumn(statusText, 0);

            StackPanel buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
            };
            buttons.Children.Add(previewButton);
            buttons.Children.Add(applyButton);
            buttons.Children.Add(closeButton);
            footer.Children.Add(buttons);
            Grid.SetColumn(buttons, 1);

            root.Children.Add(footer);
            Grid.SetRow(footer, 5);

            return root;
        }
    }
}
