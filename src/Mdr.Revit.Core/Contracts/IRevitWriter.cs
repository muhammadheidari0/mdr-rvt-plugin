using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Core.Contracts
{
    public interface IRevitWriter
    {
        SiteLogApplyResult ApplySiteLogRows(SiteLogPullResponse pullResponse);
    }
}
