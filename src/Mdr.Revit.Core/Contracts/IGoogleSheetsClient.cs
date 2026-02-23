using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Core.Contracts
{
    public interface IGoogleSheetsClient
    {
        Task<GoogleSheetReadResult> ReadRowsAsync(
            GoogleSheetSyncProfile profile,
            CancellationToken cancellationToken);

        Task<GoogleSheetWriteResult> WriteRowsAsync(
            GoogleSheetSyncProfile profile,
            System.Collections.Generic.IReadOnlyList<ScheduleSyncRow> rows,
            CancellationToken cancellationToken);
    }
}
