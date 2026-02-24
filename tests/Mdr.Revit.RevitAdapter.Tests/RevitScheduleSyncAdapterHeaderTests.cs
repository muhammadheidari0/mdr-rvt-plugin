using Mdr.Revit.RevitAdapter.Extractors;
using Xunit;

namespace Mdr.Revit.RevitAdapter.Tests
{
    public sealed class RevitScheduleSyncAdapterHeaderTests
    {
        [Fact]
        public void BuildUniqueHeaders_AppendsStableSuffixForDuplicates()
        {
            var headers = RevitScheduleSyncAdapter.BuildUniqueHeaders(
                new[] { "Area", "Area", "Area", "Level" });

            Assert.Collection(
                headers,
                x => Assert.Equal("Area", x),
                x => Assert.Equal("Area_2", x),
                x => Assert.Equal("Area_3", x),
                x => Assert.Equal("Level", x));
        }

        [Fact]
        public void BuildUniqueHeaders_NormalizesWhitespaceAndFillsEmptyHeaders()
        {
            var headers = RevitScheduleSyncAdapter.BuildUniqueHeaders(
                new[] { "  Room   Name  ", "\tRoom Name\t", " ", string.Empty });

            Assert.Collection(
                headers,
                x => Assert.Equal("Room Name", x),
                x => Assert.Equal("Room Name_2", x),
                x => Assert.Equal("COL_3", x),
                x => Assert.Equal("COL_4", x));
        }
    }
}
