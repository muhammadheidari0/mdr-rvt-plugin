using Mdr.Revit.Core.Validation;
using Xunit;

namespace Mdr.Revit.Core.Tests
{
    public sealed class SmartNumberingFormulaParserTests
    {
        [Fact]
        public void Parse_ParsesPlaceholdersAndSequenceWidth()
        {
            SmartNumberingFormulaParser parser = new SmartNumberingFormulaParser();

            var parsed = parser.Parse("{Block}{Level}-{CategoryCode}{SubcategoryCode}{Sequence:5}");

            Assert.Equal(6, parsed.Tokens.Count);
            Assert.Equal(5, parsed.SequenceWidth);
            Assert.Equal("placeholder", parsed.Tokens[0].Kind);
            Assert.Equal("Block", parsed.Tokens[0].Value);
            Assert.Equal("literal", parsed.Tokens[2].Kind);
            Assert.Equal("-", parsed.Tokens[2].Value);
            Assert.Equal("sequence", parsed.Tokens[5].Kind);
            Assert.Equal(5, parsed.Tokens[5].SequenceWidth);
        }

        [Fact]
        public void Parse_WithoutSequenceWidth_UsesFallback()
        {
            SmartNumberingFormulaParser parser = new SmartNumberingFormulaParser();

            var parsed = parser.Parse("{Block}-{Sequence}", fallbackSequenceWidth: 4);

            Assert.Equal(4, parsed.SequenceWidth);
            Assert.Equal("sequence", parsed.Tokens[2].Kind);
            Assert.Equal(4, parsed.Tokens[2].SequenceWidth);
        }

        [Fact]
        public void Parse_InvalidFormula_Throws()
        {
            SmartNumberingFormulaParser parser = new SmartNumberingFormulaParser();

            Assert.Throws<System.InvalidOperationException>(() => parser.Parse("{Block{Sequence:5}"));
        }
    }
}
