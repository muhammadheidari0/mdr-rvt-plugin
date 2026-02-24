using System;
using System.Collections.Generic;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Addin.UI
{
    public sealed class GoogleSyncWindowViewModel
    {
        private readonly List<string> _scheduleNames = new List<string>();
        private readonly List<GoogleSyncMappingItem> _columnMappings = new List<GoogleSyncMappingItem>();

        public GoogleSyncWindowViewModel()
        {
            EnsureDefaultMappings();
        }

        public IReadOnlyList<string> ScheduleNames => _scheduleNames;

        public IReadOnlyList<GoogleSyncMappingItem> ColumnMappings => _columnMappings;

        public IReadOnlyList<string> Directions { get; } = new[]
        {
            GoogleSyncDirections.Export,
            GoogleSyncDirections.Import,
        };

        public string Direction { get; set; } = GoogleSyncDirections.Export;

        public string SelectedScheduleName { get; set; } = string.Empty;

        public string SpreadsheetId { get; set; } = string.Empty;

        public string WorksheetName { get; set; } = "Sheet1";

        public string AnchorColumn { get; set; } = "MDR_UNIQUE_ID";

        public bool AuthorizeInteractively { get; set; }

        public bool PreviewOnly { get; set; }

        public void SetDefaults(string spreadsheetId, string worksheetName, string anchorColumn)
        {
            if (!string.IsNullOrWhiteSpace(spreadsheetId))
            {
                SpreadsheetId = spreadsheetId.Trim();
            }

            if (!string.IsNullOrWhiteSpace(worksheetName))
            {
                WorksheetName = worksheetName.Trim();
            }

            if (!string.IsNullOrWhiteSpace(anchorColumn))
            {
                AnchorColumn = anchorColumn.Trim();
            }
        }

        public void SetScheduleNames(IReadOnlyList<string> names)
        {
            _scheduleNames.Clear();
            if (names != null)
            {
                for (int i = 0; i < names.Count; i++)
                {
                    string value = (names[i] ?? string.Empty).Trim();
                    if (string.IsNullOrWhiteSpace(value))
                    {
                        continue;
                    }

                    if (_scheduleNames.Exists(x => string.Equals(x, value, StringComparison.OrdinalIgnoreCase)))
                    {
                        continue;
                    }

                    _scheduleNames.Add(value);
                }
            }

            if (_scheduleNames.Count == 0)
            {
                SelectedScheduleName = string.Empty;
                return;
            }

            bool hasCurrentSelection = _scheduleNames.Exists(x =>
                string.Equals(x, SelectedScheduleName, StringComparison.OrdinalIgnoreCase));
            if (!hasCurrentSelection)
            {
                SelectedScheduleName = _scheduleNames[0];
            }
        }

        public void ApplyProtectedColumns(IReadOnlyList<string> protectedColumns)
        {
            if (protectedColumns == null || protectedColumns.Count == 0)
            {
                return;
            }

            for (int i = 0; i < protectedColumns.Count; i++)
            {
                string name = (protectedColumns[i] ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                GoogleSyncMappingItem? existing = _columnMappings.Find(x =>
                    string.Equals(x.SheetColumn, name, StringComparison.OrdinalIgnoreCase));
                if (existing == null)
                {
                    _columnMappings.Add(new GoogleSyncMappingItem
                    {
                        SheetColumn = name,
                        RevitParameter = name,
                        IsEditable = false,
                    });
                    continue;
                }

                existing.IsEditable = false;
                if (string.IsNullOrWhiteSpace(existing.RevitParameter))
                {
                    existing.RevitParameter = name;
                }
            }
        }

        public void EnsureDefaultMappings()
        {
            if (_columnMappings.Count > 0)
            {
                return;
            }

            _columnMappings.Add(new GoogleSyncMappingItem
            {
                SheetColumn = "MDR_UNIQUE_ID",
                RevitParameter = "MDR_UNIQUE_ID",
                IsEditable = false,
            });

            _columnMappings.Add(new GoogleSyncMappingItem
            {
                SheetColumn = "MDR_ELEMENT_ID",
                RevitParameter = "MDR_ELEMENT_ID",
                IsEditable = false,
            });
        }

        public GoogleSyncFromAppRequest BuildRequest()
        {
            GoogleSyncFromAppRequest request = new GoogleSyncFromAppRequest
            {
                Direction = NormalizeDirection(Direction),
                ScheduleName = (SelectedScheduleName ?? string.Empty).Trim(),
                SpreadsheetId = NormalizeSpreadsheetId(SpreadsheetId),
                WorksheetName = (WorksheetName ?? string.Empty).Trim(),
                AnchorColumn = string.IsNullOrWhiteSpace(AnchorColumn) ? "MDR_UNIQUE_ID" : AnchorColumn.Trim(),
                AuthorizeInteractively = AuthorizeInteractively,
                PreviewOnly = PreviewOnly,
            };

            for (int i = 0; i < _columnMappings.Count; i++)
            {
                GoogleSyncMappingItem row = _columnMappings[i];
                if (string.IsNullOrWhiteSpace(row.SheetColumn) || string.IsNullOrWhiteSpace(row.RevitParameter))
                {
                    continue;
                }

                request.ColumnMappings.Add(new GoogleSheetColumnMapping
                {
                    SheetColumn = row.SheetColumn.Trim(),
                    RevitParameter = row.RevitParameter.Trim(),
                    IsEditable = row.IsEditable,
                });
            }

            return request;
        }

        private static string NormalizeDirection(string value)
        {
            if (string.Equals(value, GoogleSyncDirections.Import, StringComparison.OrdinalIgnoreCase))
            {
                return GoogleSyncDirections.Import;
            }

            return GoogleSyncDirections.Export;
        }

        internal static string NormalizeSpreadsheetId(string value)
        {
            string input = (value ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(input))
            {
                return string.Empty;
            }

            const string marker = "/spreadsheets/d/";
            int markerIndex = input.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
            if (markerIndex < 0)
            {
                return input;
            }

            string idPart = input.Substring(markerIndex + marker.Length);
            int endIndex = idPart.IndexOfAny(new[] { '/', '?', '#', '&' });
            if (endIndex >= 0)
            {
                idPart = idPart.Substring(0, endIndex);
            }

            return idPart.Trim();
        }
    }

    public sealed class GoogleSyncMappingItem
    {
        public string SheetColumn { get; set; } = string.Empty;

        public string RevitParameter { get; set; } = string.Empty;

        public bool IsEditable { get; set; } = true;
    }
}
