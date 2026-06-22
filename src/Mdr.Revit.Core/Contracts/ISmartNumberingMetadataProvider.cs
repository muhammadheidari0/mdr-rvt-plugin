using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Core.Contracts
{
    public interface ISmartNumberingMetadataProvider
    {
        SmartNumberingMetadata GetMetadata(SmartNumberingRule rule);
    }
}
