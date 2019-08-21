using Rubberduck.RegexAssistant.i18n;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Rubberduck.RegexAssistant.Atoms
{
    internal class Literal : IAtom
    {
        public static readonly string Pattern = @"(?<expression>\\(u[\dA-F]{4}|x[\dA-F]{2}|[0-7]{3}|[bB\(\){}\\\[\]\.+*?1-9nftvrdDwWsS])|[^()\[\]{}\\*+?^$])";
        private static readonly Regex Matcher = new Regex($"^{Pattern}$", RegexOptions.Compiled);
        private static readonly ISet<char> EscapeLiterals = new HashSet<char>();

        static Literal() {
            foreach (var escape in new[]{ '.', '+', '*', '?', '(', ')', '{', '}', '[', ']', '|', '\\' })
            {
                EscapeLiterals.Add(escape);
            }
            _escapeDescriptions.Add('d', AssistantResources.AtomDescription_Digit);
            _escapeDescriptions.Add('D', AssistantResources.AtomDescription_NonDigit);
            _escapeDescriptions.Add('b', AssistantResources.AtomDescription_WordBoundary);
            _escapeDescriptions.Add('B', AssistantResources.AtomDescription_NonWordBoundary);
            _escapeDescriptions.Add('w', AssistantResources.AtomDescription_WordCharacter);
            _escapeDescriptions.Add('W', AssistantResources.AtomDescription_NonWordCharacter);
            _escapeDescriptions.Add('s', AssistantResources.AtomDescription_Whitespace);
            _escapeDescriptions.Add('S', AssistantResources.AtomDescription_NonWhitespace);
            _escapeDescriptions.Add('n', AssistantResources.AtomDescription_Newline);
            _escapeDescriptions.Add('r', AssistantResources.AtomDescription_CarriageReturn);
            _escapeDescriptions.Add('f', AssistantResources.AtomDescription_FormFeed);
            _escapeDescriptions.Add('v', AssistantResources.AtomDescription_VTab);
            _escapeDescriptions.Add('t', AssistantResources.AtomDescription_HTab);
        }

        public Literal(string specifier, Quantifier quantifier)
        {
            if (specifier == null || quantifier == null)
            {
                throw new ArgumentNullException();
            }

            Quantifier = quantifier;
            var m = Matcher.Match(specifier);
            if (!m.Success)
            {
                throw new ArgumentException("The given specifier does not denote a Literal");
            }

            Specifier = specifier;
        }

        public string Specifier { get; }

        public Quantifier Quantifier { get; }

        private static readonly Dictionary<char, string> _escapeDescriptions = new Dictionary<char, string>();
        public string Description(bool spellOutWhitespace)
        {
            // here be dragons!
            // keep track of:
            // - escaped chars
            // - escape sequences (each having a different description)
            // - codepoint escapes (belongs into above category but kept separate)
            // - and actually boring literal matches
            if (Specifier.Length > 1)
            {
                var relevant = Specifier.Substring(1); // skip the damn Backslash at the start
                if (relevant.Length > 1) // longer sequences
                {
                    if (relevant.StartsWith("u"))
                    {
                        return string.Format(AssistantResources.AtomDescription_Literal_UnicodePoint, relevant.Substring(1)); //skip u
                    }

                    if (relevant.StartsWith("x"))
                    {
                        return string.Format(AssistantResources.AtomDescription_Literal_HexCodepoint, relevant.Substring(1)); // skip x
                    }

                    return string.Format(AssistantResources.AtomDescription_Literal_OctalCodepoint, relevant); // no format specifier to skip
                }

                if (EscapeLiterals.Contains(relevant[0]))
                {
                    return string.Format(AssistantResources.AtomDescription_Literal_EscapedLiteral, relevant);
                }

                if (char.IsDigit(relevant[0]))
                {
                    return string.Format(AssistantResources.AtomDescription_Literal_Backreference, relevant);
                }

                return _escapeDescriptions[relevant[0]];
            }

            if (Specifier.Equals("."))
            {
                return AssistantResources.AtomDescription_Dot;
            }

            return string.Format(AssistantResources.AtomDescription_Literal_ActualLiteral,
                spellOutWhitespace && WhitespaceToString.IsFullySpellingOutApplicable(Specifier, out var spelledOutWhiteSpace)
                    ? spelledOutWhiteSpace
                    : Specifier);
        }

        public override string ToString() => Specifier;
        public override bool Equals(object obj)
        {
            return obj is Literal other
                && other.Quantifier.Equals(Quantifier)
                && other.Specifier.Equals(Specifier);
        }
        public override int GetHashCode() => HashCode.Compute(Specifier, Quantifier);
    }
}
