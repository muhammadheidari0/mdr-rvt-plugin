using System.Threading;

namespace Mdr.Revit.Infra.Logging
{
    public static class CorrelationContext
    {
        private static readonly AsyncLocal<string> CurrentValue = new AsyncLocal<string>();

        public static string CurrentRunUid
        {
            get
            {
                return CurrentValue.Value ?? string.Empty;
            }

            set
            {
                CurrentValue.Value = value;
            }
        }
    }
}
