using System.Linq;
using Mdr.Revit.Addin;
using Xunit;

namespace Mdr.Revit.Addin.Tests
{
    public sealed class AppAndRibbonTests
    {
        [Fact]
        public void BuildRibbon_ReturnsExpectedCommandIds()
        {
            App app = new App();
            var commands = app.BuildRibbon();

            Assert.Equal(8, commands.Count);
            Assert.Contains(commands, x => x.Id == "mdr.login");
            Assert.Contains(commands, x => x.Id == "mdr.publishSheets");
            Assert.Contains(commands, x => x.Id == "mdr.pushSchedules");
            Assert.Contains(commands, x => x.Id == "mdr.syncSiteLogs");
            Assert.Contains(commands, x => x.Id == "mdr.googleSync");
            Assert.Contains(commands, x => x.Id == "mdr.smartNumbering");
            Assert.Contains(commands, x => x.Id == "mdr.checkUpdates");
            Assert.Contains(commands, x => x.Id == "mdr.settings");

            Assert.True(commands.All(x => !string.IsNullOrWhiteSpace(x.Title)));
        }
    }
}
