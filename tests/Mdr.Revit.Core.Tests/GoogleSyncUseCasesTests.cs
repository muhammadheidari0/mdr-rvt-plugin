using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;
using Mdr.Revit.Core.UseCases;
using Xunit;

namespace Mdr.Revit.Core.Tests
{
    public sealed class GoogleSyncUseCasesTests
    {
        [Fact]
        public async Task SyncScheduleToGoogle_ExtractsAndWritesRows()
        {
            FakeGoogleSheetsClient google = new FakeGoogleSheetsClient();
            FakeScheduleAdapter adapter = new FakeScheduleAdapter();
            adapter.Extracted.Add(new ScheduleSyncRow
            {
                AnchorUniqueId = "uid-1",
            });

            SyncScheduleToGoogleUseCase useCase = new SyncScheduleToGoogleUseCase(google, adapter);
            GoogleScheduleSyncResult result = await useCase.ExecuteAsync(
                new GoogleScheduleSyncRequest
                {
                    Direction = GoogleSyncDirections.Export,
                    ScheduleName = "MTO",
                    Profile = NewProfile(),
                },
                CancellationToken.None);

            Assert.Equal(1, result.ExportedRows);
            Assert.Equal(1, google.WriteCount);
            Assert.Equal(1, adapter.ExtractCount);
        }

        [Fact]
        public async Task SyncScheduleFromGoogle_WhenPreviewOnly_DoesNotApply()
        {
            FakeGoogleSheetsClient google = new FakeGoogleSheetsClient();
            FakeScheduleAdapter adapter = new FakeScheduleAdapter();
            google.Read.Rows.Add(new ScheduleSyncRow
            {
                AnchorUniqueId = "uid-1",
            });
            adapter.Diff.Rows.Add(new ScheduleSyncRow
            {
                AnchorUniqueId = "uid-1",
                ChangeState = ScheduleSyncStates.Modified,
            });
            adapter.Diff.ChangedRowsCount = 1;

            SyncScheduleFromGoogleUseCase useCase = new SyncScheduleFromGoogleUseCase(google, adapter);
            GoogleScheduleSyncResult result = await useCase.ExecuteAsync(
                new GoogleScheduleSyncRequest
                {
                    Direction = GoogleSyncDirections.Import,
                    Profile = NewProfile(),
                    ApplyChanges = false,
                },
                CancellationToken.None);

            Assert.Equal(1, result.DiffResult.ChangedRowsCount);
            Assert.Equal(0, adapter.ApplyCount);
        }

        private static GoogleSheetSyncProfile NewProfile()
        {
            return new GoogleSheetSyncProfile
            {
                SpreadsheetId = "sheet-id",
                WorksheetName = "Sheet1",
            };
        }

        private sealed class FakeGoogleSheetsClient : IGoogleSheetsClient
        {
            public GoogleSheetReadResult Read { get; } = new GoogleSheetReadResult();

            public int WriteCount { get; private set; }

            public Task<GoogleSheetReadResult> ReadRowsAsync(
                GoogleSheetSyncProfile profile,
                CancellationToken cancellationToken)
            {
                _ = profile;
                _ = cancellationToken;
                return Task.FromResult(Read);
            }

            public Task<GoogleSheetWriteResult> WriteRowsAsync(
                GoogleSheetSyncProfile profile,
                IReadOnlyList<ScheduleSyncRow> rows,
                CancellationToken cancellationToken)
            {
                _ = profile;
                _ = rows;
                _ = cancellationToken;
                WriteCount++;
                return Task.FromResult(new GoogleSheetWriteResult
                {
                    UpdatedRows = rows.Count,
                });
            }
        }

        private sealed class FakeScheduleAdapter : IRevitScheduleSyncAdapter
        {
            public List<ScheduleSyncRow> Extracted { get; } = new List<ScheduleSyncRow>();

            public ScheduleSyncDiffResult Diff { get; } = new ScheduleSyncDiffResult();

            public int ExtractCount { get; private set; }

            public int ApplyCount { get; private set; }

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
                _ = profile;
                return Diff;
            }

            public ScheduleSyncApplyResult ApplyDiff(ScheduleSyncDiffResult diff, GoogleSheetSyncProfile profile)
            {
                _ = diff;
                _ = profile;
                ApplyCount++;
                return new ScheduleSyncApplyResult
                {
                    AppliedCount = diff.ChangedRowsCount,
                };
            }
        }
    }
}
