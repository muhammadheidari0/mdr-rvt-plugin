using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Core.Contracts
{
    public interface IExcelWorkbookClient
    {
        Task<ExcelWorkbookReadResult> ReadRowsAsync(
            ExcelWorkbookProfile profile,
            CancellationToken cancellationToken);

        Task<ExcelWorkbookWriteResult> WriteRowsAsync(
            ExcelWorkbookProfile profile,
            IReadOnlyList<ScheduleSyncRow> rows,
            CancellationToken cancellationToken);
    }
}
