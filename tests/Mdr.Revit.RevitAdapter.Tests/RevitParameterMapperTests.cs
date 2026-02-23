using System.Collections.Generic;
using Mdr.Revit.RevitAdapter.Writers;
using Xunit;

namespace Mdr.Revit.RevitAdapter.Tests
{
    public sealed class RevitParameterMapperTests
    {
        [Fact]
        public void BuildParameterMap_CopiesInputEntries()
        {
            RevitParameterMapper mapper = new RevitParameterMapper();
            Dictionary<string, string> source = new Dictionary<string, string>
            {
                ["MDR_SYNC_KEY"] = "site_log:1:MANPOWER:1",
                ["MDR_SECTION"] = "MANPOWER",
            };

            Dictionary<string, string> result = mapper.BuildParameterMap(source);

            Assert.Equal(2, result.Count);
            Assert.Equal("MANPOWER", result["MDR_SECTION"]);
        }
    }
}
