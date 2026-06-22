using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;
using Mdr.Revit.Core.UseCases;
using Xunit;

namespace Mdr.Revit.Core.Tests
{
    public sealed class ExcelSyncUseCasesTests
    {
        [Fact]
        public async Task SyncScheduleToExcel_ExtractsAndWritesRows()
        {
            FakeExcelClient excel = new FakeExcelClient();
            FakeScheduleAdapter adapter = new FakeScheduleAdapter();
            adapter.Extracted.Add(new ScheduleSyncRow
            {
                AnchorUniqueId = "uid-1",
            });

            SyncScheduleToExcelUseCase useCase = new SyncScheduleToExcelUseCase(excel, adapter);
            ExcelScheduleSyncResult result = await useCase.ExecuteAsync(
                new ExcelScheduleSyncRequest
                {
                    Direction = GoogleSyncDirections.Export,
                    ScheduleName = "MTO",
                    Profile = NewProfile(),
                },
                CancellationToken.None);

            Assert.Equal(1, result.ExportedRows);
            Assert.Equal(1, excel.WriteCount);
            Assert.Equal(1, adapter.ExtractCount);
        }

        [Fact]
        public async Task SyncScheduleFromExcel_UsesInstanceOnlyDiffAndSkipsApplyWhenPreview()
        {
            FakeExcelClient excel = new FakeExcelClient();
            FakeScheduleAdapter adapter = new FakeScheduleAdapter();
            excel.Read.Rows.Add(new ScheduleSyncRow
            {
                AnchorUniqueId = "uid-1",
            });
            adapter.Diff.ChangedRowsCount = 1;
            adapter.Diff.Rows.Add(new ScheduleSyncRow
            {
                AnchorUniqueId = "uid-1",
                ChangeState = ScheduleSyncStates.Modified,
            });

            SyncScheduleFromExcelUseCase useCase = new SyncScheduleFromExcelUseCase(excel, adapter);
            ExcelScheduleSyncResult result = await useCase.ExecuteAsync(
                new ExcelScheduleSyncRequest
                {
                    Direction = GoogleSyncDirections.Import,
                    Profile = NewProfile(),
                    ApplyChanges = false,
                },
                CancellationToken.None);

            Assert.Equal(1, result.DiffResult.ChangedRowsCount);
            Assert.True(adapter.LastDiffProfile?.InstanceOnlyWrites);
            Assert.Equal(0, adapter.ApplyCount);
        }

        private static ExcelWorkbookProfile NewProfile()
        {
            ExcelWorkbookProfile profile = new ExcelWorkbookProfile
            {
                FilePath = "C:\\temp\\mto.xlsx",
                WorksheetName = "MTO",
                AnchorColumn = "MDR_UNIQUE_ID",
            };
            profile.ColumnMappings.Add(new GoogleSheetColumnMapping
            {
                SheetColumn = "Comments",
                RevitParameter = "Comments",
                IsEditable = true,
            });
            return profile;
        }

        private sealed class FakeExcelClient : IExcelWorkbookClient
        {
            public ExcelWorkbookReadResult Read { get; } = new ExcelWorkbookReadResult();

            public int WriteCount { get; private set; }

            public Task<ExcelWorkbookReadResult> ReadRowsAsync(
                ExcelWorkbookProfile profile,
                CancellationToken cancellationToken)
            {
                _ = profile;
                _ = cancellationToken;
                return Task.FromResult(Read);
            }

            public Task<ExcelWorkbookWriteResult> WriteRowsAsync(
                ExcelWorkbookProfile profile,
                IReadOnlyList<ScheduleSyncRow> rows,
                CancellationToken cancellationToken)
            {
                _ = profile;
                _ = rows;
                _ = cancellationToken;
                WriteCount++;
                return Task.FromResult(new ExcelWorkbookWriteResult
                {
                    UpdatedRows = rows.Count,
                });
            }
        }

        private sealed class FakeScheduleAdapter : IRevitScheduleSyncAdapter
        {
            public List<ScheduleSyncRow> Extracted { get; } = new List<ScheduleSyncRow>();

            public ScheduleSyncDiffResult Diff { get; } = new ScheduleSyncDiffResult();

            public GoogleSheetSyncProfile? LastDiffProfile { get; private set; }

            public int ExtractCount { get; private set; }

            public int ApplyCount { get; private set; }

            public IReadOnlyList<string> GetAvailableScheduleNames()
            {
                return new[] { "MTO" };
            }

            public IReadOnlyList<GoogleSheetColumnMapping> GetScheduleColumnMappings(string scheduleName)
            {
                _ = scheduleName;
                return new[] { new GoogleSheetColumnMapping { SheetColumn = "Comments", RevitParameter = "Comments" } };
            }

            public IReadOnlyList<ScheduleSyncRow> ExtractRows(string scheduleName, GoogleSheetSyncProfile profile)
            {
                _ = scheduleName;
                _ = profile;
                ExtractCount++;
                return Extracted;
            }

            public ScheduleSyncDiffResult BuildDiff(
                IReadOnlyList<ScheduleSyncRow> incomingRows,
                GoogleSheetSyncProfile profile)
            {
                _ = incomingRows;
                LastDiffProfile = profile;
                return Diff;
            }

            public ScheduleSyncApplyResult ApplyDiff(ScheduleSyncDiffResult diff, GoogleSheetSyncProfile profile)
            {
                _ = diff;
                _ = profile;
                ApplyCount++;
                return new ScheduleSyncApplyResult();
            }
        }
    }
}
