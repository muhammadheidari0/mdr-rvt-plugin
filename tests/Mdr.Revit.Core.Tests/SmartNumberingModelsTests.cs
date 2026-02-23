using Mdr.Revit.Core.Models;
using Xunit;

namespace Mdr.Revit.Core.Tests
{
    public sealed class SmartNumberingModelsTests
    {
        [Fact]
        public void PreviewItem_ValueAlias_ReadsAndWritesProposedValue()
        {
            SmartNumberingPreviewItem item = new SmartNumberingPreviewItem
            {
                ProposedValue = "AL03WS00001",
            };

            Assert.Equal("AL03WS00001", item.Value);

            item.Value = "AL03WS00002";
            Assert.Equal("AL03WS00002", item.ProposedValue);
        }

        [Fact]
        public void Result_DefaultsToAtomicMode()
        {
            SmartNumberingResult result = new SmartNumberingResult();

            Assert.True(result.IsAtomicRollback);
            Assert.False(result.WasRolledBack);
            Assert.Equal(string.Empty, result.FatalErrorCode);
        }
    }
}
