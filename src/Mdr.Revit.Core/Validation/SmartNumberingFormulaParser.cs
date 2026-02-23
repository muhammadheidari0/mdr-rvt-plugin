using System;
using System.Globalization;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Core.Validation
{
    public sealed class SmartNumberingFormulaParser
    {
        public SmartNumberingFormula Parse(string formulaText, int fallbackSequenceWidth = 5)
        {
            if (string.IsNullOrWhiteSpace(formulaText))
            {
                throw new InvalidOperationException("Smart numbering formula is required.");
            }

            SmartNumberingFormula formula = new SmartNumberingFormula
            {
                SequenceWidth = fallbackSequenceWidth <= 0 ? 5 : fallbackSequenceWidth,
            };

            string input = formulaText.Trim();
            int index = 0;
            while (index < input.Length)
            {
                int open = input.IndexOf('{', index);
                if (open < 0)
                {
                    AddLiteralToken(formula, input.Substring(index));
                    break;
                }

                if (open > index)
                {
                    AddLiteralToken(formula, input.Substring(index, open - index));
                }

                int close = input.IndexOf('}', open + 1);
                if (close < 0)
                {
                    throw new InvalidOperationException("Smart numbering formula has unmatched '{'.");
                }

                string content = input.Substring(open + 1, close - open - 1).Trim();
                if (string.IsNullOrWhiteSpace(content))
                {
                    throw new InvalidOperationException("Smart numbering placeholder cannot be empty.");
                }

                if (content.IndexOf('{') >= 0 || content.IndexOf('}') >= 0)
                {
                    throw new InvalidOperationException("Smart numbering placeholder has invalid nested braces.");
                }

                if (content.StartsWith("Sequence", StringComparison.OrdinalIgnoreCase))
                {
                    int width = ParseSequenceWidth(content, formula.SequenceWidth);
                    formula.SequenceWidth = width;
                    formula.Tokens.Add(new SmartNumberingToken
                    {
                        Kind = "sequence",
                        Value = "Sequence",
                        SequenceWidth = width,
                    });
                }
                else
                {
                    formula.Tokens.Add(new SmartNumberingToken
                    {
                        Kind = "placeholder",
                        Value = content,
                    });
                }

                index = close + 1;
            }

            if (formula.Tokens.Count == 0)
            {
                throw new InvalidOperationException("Smart numbering formula cannot be empty.");
            }

            return formula;
        }

        private static void AddLiteralToken(SmartNumberingFormula formula, string literal)
        {
            if (string.IsNullOrEmpty(literal))
            {
                return;
            }

            formula.Tokens.Add(new SmartNumberingToken
            {
                Kind = "literal",
                Value = literal,
            });
        }

        private static int ParseSequenceWidth(string token, int fallback)
        {
            int width = fallback > 0 ? fallback : 5;
            int idx = token.IndexOf(':');
            if (idx < 0)
            {
                return width;
            }

            string suffix = token.Substring(idx + 1).Trim();
            if (string.IsNullOrWhiteSpace(suffix))
            {
                return width;
            }

            if (!int.TryParse(suffix, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed))
            {
                throw new InvalidOperationException("Sequence width must be an integer.");
            }

            if (parsed <= 0 || parsed > 12)
            {
                throw new InvalidOperationException("Sequence width must be between 1 and 12.");
            }

            return parsed;
        }
    }
}
