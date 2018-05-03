using Rubberduck.RegexAssistant.i18n;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Rubberduck.RegexAssistant
{
    public interface IAtom : IDescribable
    {
        Quantifier Quantifier { get; }
        string Specifier { get; }
    }
    
    internal class CharacterClass : IAtom
    {
        public static readonly string Pattern = @"(?<!\\)\[(?<expression>.*?)(?<!\\)\]";
        private static readonly Regex Matcher = new Regex($"^{Pattern}$", RegexOptions.Compiled);
        
        public bool InverseMatching { get; }
        public IList<string> CharacterSpecifiers { get; }

        public CharacterClass(string specifier, Quantifier quantifier)
        {
            if (specifier == null || quantifier == null)
            {
                throw new ArgumentNullException();
            }

            Quantifier = quantifier;
            var m = Matcher.Match(specifier);
            if (!m.Success)
            {
                throw new ArgumentException("The given specifier does not denote a character class");
            }
            Specifier = specifier;
            var actualSpecifier = m.Groups["expression"].Value;
            InverseMatching = actualSpecifier.StartsWith("^");
            CharacterSpecifiers= ExtractCharacterSpecifiers(InverseMatching ? actualSpecifier.Substring(1) : actualSpecifier);
        }

        public string Specifier { get; }

        public Quantifier Quantifier { get; }

        private static readonly Regex CharacterRanges = new Regex(@"(\\[dDwWsS]|(\\[ntfvr]|\\([0-7]{3}|x[\dA-F]{2}|u[\dA-F]{4}|[\\\.\[\]])|.)(-(\\[ntfvr]|\\([0-7]{3}|x[A-F]{2}|u[\dA-F]{4}|[\.\\\[\]])|.))?)", RegexOptions.Compiled);
        private IList<string> ExtractCharacterSpecifiers(string characterClass)
        {
            var specifiers = CharacterRanges.Matches(characterClass);
            var result = new List<string>();
            
            foreach (Match specifier in specifiers)
            {
                if (specifier.Value.Contains("\\"))
                {
                    if (specifier.Value.EndsWith("-\\"))
                    {
                        throw new ArgumentException("Character Ranges that have incorrectly escaped characters as target are not allowed");
                    }

                    if (specifier.Value.Length == 1)
                    {
                        // Something's bork with the Pattern. For now we skip this it shouldn't affect anyone
                        continue;
                    }
                }
                result.Add(specifier.Value);
            }
            return result;
        }

        public string Description => string.Format(InverseMatching 
                ? AssistantResources.AtomDescription_CharacterClass_Inverted 
                : AssistantResources.AtomDescription_CharacterClass
            , HumanReadableClass());

        private string HumanReadableClass()
        {
            return string.Join(", ", CharacterSpecifiers); // join last with and?
        }

        public override bool Equals(object obj)
        {
            return obj is CharacterClass other 
                && other.Quantifier.Equals(Quantifier)
                && other.Specifier.Equals(Specifier);
        }

        public override int GetHashCode()
        {
            return Specifier.GetHashCode();
        }
    }

    public class Group : IAtom
    {
        public static readonly string Pattern = @"(?<!\\)\((?<expression>.*(?<!\\))\)";
        private static readonly Regex Matcher = new Regex($"^{Pattern}$", RegexOptions.Compiled);

        public Group(string specifier, Quantifier quantifier) {
            if (specifier == null || quantifier == null)
            {
                throw new ArgumentNullException();
            }

            Quantifier = quantifier;
            var m = Matcher.Match(specifier);
            if (!m.Success)
            {
                throw new ArgumentException("The given specifier does not denote a Group");
            }
            Subexpression = RegularExpression.Parse(m.Groups["expression"].Value);
            Specifier = specifier;
        }

        public IRegularExpression Subexpression { get; }

        public Quantifier Quantifier { get; }

        public string Specifier { get; }

        public string Description => string.Format(AssistantResources.AtomDescription_Group, Specifier);

        public override bool Equals(object obj)
        {
            return obj is Group other
                && other.Quantifier.Equals(Quantifier)
                && other.Specifier.Equals(Specifier);
        }

        public override int GetHashCode()
        {
            return Specifier.GetHashCode();
        }
    }

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
        public string Description
        {
            get
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
                        else if (relevant.StartsWith("x"))
                        {
                            return string.Format(AssistantResources.AtomDescription_Literal_HexCodepoint, relevant.Substring(1)); // skip x
                        }
                        else
                        {
                            return string.Format(AssistantResources.AtomDescription_Literal_OctalCodepoint, relevant); // no format specifier to skip
                        }
                    }
                    else if (EscapeLiterals.Contains(relevant[0]))
                    {
                        return string.Format(AssistantResources.AtomDescription_Literal_EscapedLiteral, relevant);
                    }
                    else if (char.IsDigit(relevant[0]))
                    {
                        return string.Format(AssistantResources.AtomDescription_Literal_Backreference, relevant);
                    }
                    else
                    {
                        return _escapeDescriptions[relevant[0]];
                    }
                }

                return Specifier.Equals(".") 
                    ? AssistantResources.AtomDescription_Dot 
                    : string.Format(AssistantResources.AtomDescription_Literal_ActualLiteral, Specifier);
            }
        }

        public override bool Equals(object obj)
        {
            return obj is Literal other
                && other.Quantifier.Equals(Quantifier)
                && other.Specifier.Equals(Specifier);
        }

        public override int GetHashCode()
        {
            return Specifier.GetHashCode();
        }
    }
}
