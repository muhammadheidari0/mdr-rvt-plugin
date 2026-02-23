using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Core.Contracts
{
    public interface ISmartNumberingEngine
    {
        SmartNumberingResult Apply(SmartNumberingRule rule, bool previewOnly);
    }
}
