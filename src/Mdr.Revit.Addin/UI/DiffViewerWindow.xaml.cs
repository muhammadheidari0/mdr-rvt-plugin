using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
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

        public bool ShowDialog(bool allowApply)
        {
            Window window = new Window
            {
                Title = "Sync Preview",
                Width = 900,
                Height = 540,
                MinWidth = 720,
                MinHeight = 420,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
            };

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
                Text = "Review changes before applying to Revit",
            };
            root.Children.Add(title);
            Grid.SetRow(title, 0);

            TextBlock summary = new TextBlock
            {
                Margin = new Thickness(0, 0, 0, 12),
                Text = "Changed rows: " + Diff.ChangedRowsCount + " | Error rows: " + Diff.ErrorRowsCount,
            };
            root.Children.Add(summary);
            Grid.SetRow(summary, 1);

            List<DiffPreviewRow> previewRows = BuildPreviewRows(Diff);
            DataGrid grid = BuildGrid(previewRows);
            root.Children.Add(grid);
            Grid.SetRow(grid, 2);

            StackPanel buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 12, 0, 0),
            };

            bool canApply = allowApply && Diff.ChangedRowsCount > 0;
            Button applyButton = new Button
            {
                Width = 90,
                Margin = new Thickness(0, 0, 8, 0),
                Content = "Apply",
                IsEnabled = canApply,
                IsDefault = canApply,
            };
            applyButton.Click += (_, _) =>
            {
                window.DialogResult = true;
                window.Close();
            };

            Button closeButton = new Button
            {
                Width = 90,
                Content = canApply ? "Close" : "OK",
                IsCancel = true,
            };
            closeButton.Click += (_, _) =>
            {
                window.DialogResult = false;
                window.Close();
            };

            buttons.Children.Add(applyButton);
            buttons.Children.Add(closeButton);

            root.Children.Add(buttons);
            Grid.SetRow(buttons, 3);

            window.Content = root;
            bool? result = window.ShowDialog();
            return result == true;
        }

        private static DataGrid BuildGrid(IReadOnlyList<DiffPreviewRow> rows)
        {
            DataGrid grid = new DataGrid
            {
                AutoGenerateColumns = false,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                ItemsSource = rows,
            };

            grid.Columns.Add(new DataGridTextColumn
            {
                Header = "Anchor",
                Width = 260,
                Binding = new Binding(nameof(DiffPreviewRow.Anchor)),
            });
            grid.Columns.Add(new DataGridTextColumn
            {
                Header = "State",
                Width = 120,
                Binding = new Binding(nameof(DiffPreviewRow.State)),
            });
            grid.Columns.Add(new DataGridTextColumn
            {
                Header = "Details",
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                Binding = new Binding(nameof(DiffPreviewRow.Details)),
            });

            return grid;
        }

        private static List<DiffPreviewRow> BuildPreviewRows(ScheduleSyncDiffResult diff)
        {
            List<DiffPreviewRow> rows = new List<DiffPreviewRow>();
            if (diff?.Rows == null || diff.Rows.Count == 0)
            {
                return rows;
            }

            for (int i = 0; i < diff.Rows.Count; i++)
            {
                ScheduleSyncRow row = diff.Rows[i];
                if (!string.Equals(row.ChangeState, ScheduleSyncStates.Modified, StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(row.ChangeState, ScheduleSyncStates.Error, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string details;
                if (row.Errors.Count > 0)
                {
                    details = string.Join(
                        "; ",
                        row.Errors.Select(x => (x.Code ?? string.Empty) + ": " + (x.Message ?? string.Empty)));
                }
                else
                {
                    details = "Mapped editable fields will be written to Revit.";
                }

                rows.Add(new DiffPreviewRow
                {
                    Anchor = string.IsNullOrWhiteSpace(row.AnchorUniqueId)
                        ? (row.ElementId ?? string.Empty)
                        : row.AnchorUniqueId,
                    State = row.ChangeState ?? string.Empty,
                    Details = details,
                });
            }

            return rows;
        }
    }

    public sealed class DiffPreviewRow
    {
        public string Anchor { get; set; } = string.Empty;

        public string State { get; set; } = string.Empty;

        public string Details { get; set; } = string.Empty;
    }
}
