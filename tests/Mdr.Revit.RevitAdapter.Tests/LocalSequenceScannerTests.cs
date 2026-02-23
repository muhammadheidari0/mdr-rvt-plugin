using Mdr.Revit.RevitAdapter.Helpers;
using Xunit;

namespace Mdr.Revit.RevitAdapter.Tests
{
    public sealed class LocalSequenceScannerTests
    {
        [Theory]
        [InlineData("AL03-WS00001", "AL03-WS", 1)]
        [InlineData("AL03-WS00025", "AL03-WS", 25)]
        [InlineData("ZZ-XXAB12", "ZZ-XX", 12)]
        [InlineData("NO_MATCH_0001", "AL03-WS", 0)]
        public void ExtractSequence_ParsesTrailingDigits(string value, string prefix, int expected)
        {
            LocalSequenceScanner scanner = new LocalSequenceScanner();
            int actual = scanner.ExtractSequence(value, prefix);
            Assert.Equal(expected, actual);
        }
    }
}
