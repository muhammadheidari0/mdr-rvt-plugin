using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Autodesk.Revit.UI;
using Mdr.Revit.Client.Excel;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;
using Mdr.Revit.Core.UseCases;
using Mdr.Revit.Infra.Logging;
using Mdr.Revit.RevitAdapter.Extractors;

namespace Mdr.Revit.Addin.Commands
{
    public sealed class ExcelSyncCommand
    {
        private readonly IRevitScheduleSyncAdapter _revitScheduleSyncAdapter;
        private readonly IExcelWorkbookClient _excelWorkbookClient;
        private readonly PluginLogger _logger;

        public ExcelSyncCommand()
            : this(
                new NullScheduleSyncAdapter(),
                new ExcelWorkbookClient(),
                new PluginLogger(DefaultLogDirectory()))
        {
        }

        public ExcelSyncCommand(UIDocument uiDocument)
            : this(
                RevitApiExtractors.CreateScheduleSyncAdapter(uiDocument),
                new ExcelWorkbookClient(),
                new PluginLogger(DefaultLogDirectory()))
        {
        }

        internal ExcelSyncCommand(
            IRevitScheduleSyncAdapter revitScheduleSyncAdapter,
            IExcelWorkbookClient excelWorkbookClient,
            PluginLogger logger)
        {
            _revitScheduleSyncAdapter = revitScheduleSyncAdapter ?? throw new ArgumentNullException(nameof(revitScheduleSyncAdapter));
            _excelWorkbookClient = excelWorkbookClient ?? throw new ArgumentNullException(nameof(excelWorkbookClient));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        public string Id => "mdr.excelSync";

        public IReadOnlyList<string> GetAvailableSchedules()
        {
            return _revitScheduleSyncAdapter.GetAvailableScheduleNames();
        }

        public IReadOnlyList<GoogleSheetColumnMapping> GetScheduleColumnMappings(string scheduleName)
        {
            return _revitScheduleSyncAdapter.GetScheduleColumnMappings(scheduleName ?? string.Empty);
        }

        public async Task<ExcelScheduleSyncResult> ExecuteAsync(
            ExcelSyncCommandRequest request,
            CancellationToken cancellationToken)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            NormalizeRequest(request);
            ValidateRequest(request);

            ExcelWorkbookProfile profile = BuildProfile(request);
            ExcelScheduleSyncRequest syncRequest = new ExcelScheduleSyncRequest
            {
                Direction = request.Direction,
                ScheduleName = request.ScheduleName ?? string.Empty,
                Profile = profile,
                ApplyChanges = !request.PreviewOnly,
            };

            if (string.Equals(request.Direction, GoogleSyncDirections.Export, StringComparison.OrdinalIgnoreCase))
            {
                SyncScheduleToExcelUseCase useCase = new SyncScheduleToExcelUseCase(
                    _excelWorkbookClient,
                    _revitScheduleSyncAdapter);
                ExcelScheduleSyncResult result = await useCase.ExecuteAsync(syncRequest, cancellationToken).ConfigureAwait(false);
                _logger.Info(
                    "Excel export completed rows=" + result.ExportedRows +
                    " skipped=" + result.SkippedRows +
                    " file=" + profile.FilePath);
                return result;
            }

            SyncScheduleFromExcelUseCase importUseCase = new SyncScheduleFromExcelUseCase(
                _excelWorkbookClient,
                _revitScheduleSyncAdapter);
            ExcelScheduleSyncResult importResult = await importUseCase.ExecuteAsync(syncRequest, cancellationToken).ConfigureAwait(false);
            _logger.Info(
                "Excel import completed changed=" + importResult.DiffResult.ChangedRowsCount +
                " errors=" + importResult.DiffResult.ErrorRowsCount +
                " applied=" + importResult.ApplyResult.AppliedCount);
            return importResult;
        }

        private static ExcelWorkbookProfile BuildProfile(ExcelSyncCommandRequest request)
        {
            HashSet<string> protectedColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "MDR_UNIQUE_ID",
                "MDR_ELEMENT_ID",
            };

            for (int i = 0; i < request.ProtectedColumns.Count; i++)
            {
                string value = (request.ProtectedColumns[i] ?? string.Empty).Trim();
                if (!string.IsNullOrWhiteSpace(value))
                {
                    protectedColumns.Add(value);
                }
            }

            ExcelWorkbookProfile profile = new ExcelWorkbookProfile
            {
                FilePath = request.FilePath,
                WorksheetName = request.WorksheetName,
                AnchorColumn = request.AnchorColumn,
            };

            HashSet<string> mappedColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < request.ColumnMappings.Count; i++)
            {
                GoogleSheetColumnMapping? mapping = request.ColumnMappings[i];
                if (mapping == null)
                {
                    continue;
                }

                string excelColumn = (mapping.SheetColumn ?? string.Empty).Trim();
                string revitParameter = (mapping.RevitParameter ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(excelColumn) || string.IsNullOrWhiteSpace(revitParameter))
                {
                    continue;
                }

                GoogleSheetColumnMapping normalized = new GoogleSheetColumnMapping
                {
                    SheetColumn = excelColumn,
                    RevitParameter = revitParameter,
                    IsEditable = mapping.IsEditable && !protectedColumns.Contains(excelColumn),
                };

                if (mappedColumns.Contains(excelColumn))
                {
                    for (int m = 0; m < profile.ColumnMappings.Count; m++)
                    {
                        if (string.Equals(profile.ColumnMappings[m].SheetColumn, excelColumn, StringComparison.OrdinalIgnoreCase))
                        {
                            profile.ColumnMappings[m] = normalized;
                            break;
                        }
                    }
                }
                else
                {
                    profile.ColumnMappings.Add(normalized);
                    mappedColumns.Add(excelColumn);
                }
            }

            foreach (string systemColumn in new[] { "MDR_UNIQUE_ID", "MDR_ELEMENT_ID" })
            {
                if (mappedColumns.Contains(systemColumn))
                {
                    for (int i = 0; i < profile.ColumnMappings.Count; i++)
                    {
                        if (string.Equals(profile.ColumnMappings[i].SheetColumn, systemColumn, StringComparison.OrdinalIgnoreCase))
                        {
                            profile.ColumnMappings[i].IsEditable = false;
                            break;
                        }
                    }
                }
                else
                {
                    profile.ColumnMappings.Add(new GoogleSheetColumnMapping
                    {
                        SheetColumn = systemColumn,
                        RevitParameter = systemColumn,
                        IsEditable = false,
                    });
                }
            }

            foreach (string protectedColumn in protectedColumns)
            {
                profile.ProtectedColumns.Add(protectedColumn);
            }

            return profile;
        }

        private static void NormalizeRequest(ExcelSyncCommandRequest request)
        {
            request.Direction = string.Equals(request.Direction, GoogleSyncDirections.Import, StringComparison.OrdinalIgnoreCase)
                ? GoogleSyncDirections.Import
                : GoogleSyncDirections.Export;
            request.FilePath = (request.FilePath ?? string.Empty).Trim();
            request.WorksheetName = (request.WorksheetName ?? string.Empty).Trim();
            request.AnchorColumn = string.IsNullOrWhiteSpace(request.AnchorColumn)
                ? "MDR_UNIQUE_ID"
                : request.AnchorColumn.Trim();
        }

        private static void ValidateRequest(ExcelSyncCommandRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.FilePath))
            {
                throw new InvalidOperationException("Excel file path is required.");
            }

            if (!request.FilePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("Excel sync supports .xlsx files only.");
            }

            if (string.IsNullOrWhiteSpace(request.WorksheetName))
            {
                throw new InvalidOperationException("WorksheetName is required.");
            }
        }

        private static string DefaultLogDirectory()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "MDR",
                "RevitPlugin",
                "logs");
        }
    }

    public sealed class ExcelSyncCommandRequest
    {
        public string Direction { get; set; } = GoogleSyncDirections.Export;

        public string ScheduleName { get; set; } = string.Empty;

        public string FilePath { get; set; } = string.Empty;

        public string WorksheetName { get; set; } = string.Empty;

        public string AnchorColumn { get; set; } = "MDR_UNIQUE_ID";

        public bool PreviewOnly { get; set; }

        public List<GoogleSheetColumnMapping> ColumnMappings { get; } = new List<GoogleSheetColumnMapping>();

        public List<string> ProtectedColumns { get; } = new List<string>();
    }
}
