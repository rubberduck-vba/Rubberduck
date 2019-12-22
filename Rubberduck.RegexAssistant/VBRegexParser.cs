using Rubberduck.RegexAssistant.Atoms;
using RdGroup = Rubberduck.RegexAssistant.Atoms.Group;
using Rubberduck.RegexAssistant.Expressions;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Rubberduck.RegexAssistant
{
    internal class VBRegexParser
    {
        private static readonly Regex LITERAL_PATTERN = new Regex("^" + Literal.Pattern);
        private static readonly Regex QUANTIFIER_PATTERN = new Regex("^" + Quantifier.Pattern);

        public static IRegularExpression Parse(string specifier, bool spellOutWhiteSpace = false)
        {
            if (specifier == null) throw new ArgumentNullException(nameof(specifier));

            var subexpressions = new List<IRegularExpression>();
            var concatenation = new List<IRegularExpression>();
            while (specifier != string.Empty)
            {
                if (specifier.StartsWith("|"))
                {
                    subexpressions.Add(concatenation.Count == 1 ? concatenation[0] : new ConcatenatedExpression(concatenation));
                    concatenation.Clear();
                    specifier = specifier.Substring(1);
                    continue;
                }
                if (specifier.StartsWith("("))
                {
                    var expressionBody = DescendGroup(specifier);
                    if (expressionBody.Length != 0)
                    {
                        var quantifier = GetQuantifier(specifier, expressionBody.Length);
                        var expression = Parse(expressionBody.Substring(1, expressionBody.Length - 2));
                        concatenation.Add(new SingleAtomExpression(new RdGroup(expression, expressionBody, new Quantifier(quantifier))));
                        specifier = specifier.Substring(expressionBody.Length + quantifier.Length);
                        continue;
                    }
                }
                if (specifier.StartsWith("["))
                {
                    var expressionBody = DescendClass(specifier);
                    if (expressionBody.Length != 0)
                    {
                        var quantifier = GetQuantifier(specifier, expressionBody.Length);
                        concatenation.Add(new SingleAtomExpression(new CharacterClass(expressionBody, new Quantifier(quantifier), spellOutWhiteSpace)));
                        specifier = specifier.Substring(expressionBody.Length + quantifier.Length);
                        continue;
                    }
                }
                // finally attempt to parse as literal. If that fails, we need to consume the character as an error expression
                {
                    var expressionBody = DescendLiteral(specifier);
                    if (expressionBody.Length == 0)
                    {
                        // well, this is an error
                        concatenation.Add(new ErrorExpression(string.Empty + specifier[0]));
                        specifier = specifier.Substring(1);
                        continue;
                    }
                    var quantifier = GetQuantifier(specifier, expressionBody.Length);
                    concatenation.Add(new SingleAtomExpression(new Literal(expressionBody, new Quantifier(quantifier))));
                    specifier = specifier.Substring(expressionBody.Length + quantifier.Length);
                    continue;
                }
            }
            
            if (subexpressions.Count != 0)
            {
                subexpressions.Add(concatenation.Count == 1 ? concatenation[0] : new ConcatenatedExpression(concatenation));
                return new AlternativesExpression(subexpressions);
            }
            return concatenation.Count == 1 ? concatenation[0] : new ConcatenatedExpression(concatenation);
        }

        private static string DescendLiteral(string specifier)
        {
            var matcher = LITERAL_PATTERN.Match(specifier);
            return matcher.Success 
                ? matcher.Groups["expression"].Value
                : string.Empty;
        }

        private static string DescendClass(string specifier)
        {
            return DescendExpression(specifier, '[', ']');
        }

        private static string GetQuantifier(string specifier, int length)
        {
            var operationalSubstring = specifier.Substring(length);
            var matcher = QUANTIFIER_PATTERN.Match(operationalSubstring);
            return matcher.Success
                ? matcher.Groups["quantifier"].Value
                : string.Empty;
        }

        private static string DescendGroup(string specifier)
        {
            return DescendExpression(specifier, '(', ')');
        }

        private static string DescendExpression(string specifier, char opening, char closing)
        {
            var length = 0;
            var openingCount = 0;
            var escapeToggle = false;
            foreach (var digit in specifier)
            {
                if (digit == opening && !escapeToggle)
                {
                    openingCount++;
                    escapeToggle = false;
                }
                if (digit == closing && !escapeToggle)
                {
                    openingCount--;
                    escapeToggle = false;
                    if (openingCount <= 0)
                    {
                        return openingCount == 0 ? specifier.Substring(0, length + 1) : string.Empty;
                    }
                }
                if (digit == '\\' || escapeToggle)
                {
                    escapeToggle = !escapeToggle;
                }
                length++;
            }
            return string.Empty;
        }
    }
}
