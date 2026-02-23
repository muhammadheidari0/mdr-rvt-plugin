using Mdr.Revit.Core.Validation;
using Xunit;

namespace Mdr.Revit.Core.Tests
{
    public sealed class SemanticVersionComparerTests
    {
        [Theory]
        [InlineData("1.2.0", "1.1.9", true)]
        [InlineData("v1.2.0", "1.2.0", false)]
        [InlineData("1.2.1", "1.2.1", false)]
        [InlineData("2.0.0-beta", "1.9.9", true)]
        public void IsGreater_HandlesCommonVersionFormats(string candidate, string current, bool expected)
        {
            bool actual = SemanticVersionComparer.IsGreater(candidate, current);
            Assert.Equal(expected, actual);
        }
    }
}
