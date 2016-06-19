using Rubberduck.RegexAssistant.i18n;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Rubberduck.RegexAssistant
{
    public interface Atom : IDescribable
    {

    }

    public class CharacterClass : Atom
    {
        public static readonly string Pattern = @"(?<!\\)\[(?<expression>.*?)(?<!\\)\]";
        private static readonly Regex Matcher = new Regex("^" + Pattern + "$");

        public bool InverseMatching { get; }
        public IList<string> CharacterSpecifiers { get; }

        public CharacterClass(string specifier)
        {
            Match m = Matcher.Match(specifier);
            if (!m.Success)
            {
                throw new ArgumentException("The given specifier does not denote a character class");
            }
            string actualSpecifier = m.Groups["expression"].Value;
            InverseMatching = actualSpecifier.StartsWith("^");
            CharacterSpecifiers = new List<string>();

            ExtractCharacterSpecifiers(InverseMatching ? actualSpecifier.Substring(1) : actualSpecifier);
        }

        private static readonly Regex CharacterRanges = new Regex(@"(\\[dDwWsS]|(\\[ntfvr]|\\([0-7]{3}|x[\dA-F]{2}|u[\dA-F]{4}|[\\\.\[\]])|.)(-(\\[ntfvr]|\\([0-7]{3}|x[A-F]{2}|u[\dA-F]{4}|[\.\\\[\]])|.))?)");
        private void ExtractCharacterSpecifiers(string characterClass)
        {
            MatchCollection specifiers = CharacterRanges.Matches(characterClass);
            
            foreach (Match specifier in specifiers)
            {
                if (specifier.Value.Contains("\\"))
                {
                    if (specifier.Value.EndsWith("-\\"))
                    {
                        // BOOM!
                        throw new ArgumentException("Character Ranges that have incorrectly escaped characters as target are not allowed");
                    }
                    else if (specifier.Value.Length == 1)
                    {
                        // fun... we somehow got to grab a single backslash. Pattern is probably broken
                        // alas for simplicity we just skip the incorrect backslash
                        // TODO: make a warning from this.. how?? no idea
                        continue;
                    }
                }
                CharacterSpecifiers.Add(specifier.Value);
            }
        }

        public string Description
        {
            get
            {
                return string.Format(InverseMatching 
                    ? AssistantResources.AtomDescription_CharacterClass_Inverted 
                    : AssistantResources.AtomDescription_CharacterClass
                    , HumanReadableClass());
            }
        }

        private string HumanReadableClass()
        {
            return string.Join(", ", CharacterSpecifiers); // join last with and?
        }

        public bool TryMatch(ref string text)
        {
            throw new NotImplementedException();
        }
    }

    class Group : Atom
    {
        public static readonly string Pattern = @"(?<!\\)\((?<expression>.*)(?<!\\)\)";
        private static readonly Regex Matcher = new Regex("^" + Pattern + "$");

        private readonly IRegularExpression subexpression;
        private readonly string specifier;

        public Group(string specifier) {
            Match m = Matcher.Match(specifier);
            if (!m.Success)
            {
                throw new ArgumentException("The given specifier does not denote a Group");
            }
            subexpression = RegularExpression.Parse(m.Groups["expression"].Value);
            this.specifier = specifier;
        }

        public string Description
        {
            get
            {
                return string.Format(AssistantResources.AtomDescription_Group, specifier) + "\r\n" + subexpression.Description;
            }
        }

        public bool TryMatch(ref string text)
        {
            throw new NotImplementedException();
        }
    }

    class Literal : Atom
    {
        public static readonly string Pattern = @"\(?:[bB(){}\\\[\]\.+*?\dnftvrdDwWsS]|u[\dA-F]{4}|x[\dA-F]{2}|[0-7]{3})|.";
        private static readonly Regex Matcher = new Regex("^" + Pattern + "$");
        private static readonly ISet<char> EscapeLiterals = new HashSet<char>();
        private readonly string specifier;

        static Literal() {
            foreach (char escape in new char[]{ '.', '+', '*', '?', '(', ')', '{', '}', '[', ']', '|', '\\' })
            {
                EscapeLiterals.Add(escape);
            }
        }


        public Literal(string specifier)
        {
            Match m = Matcher.Match(specifier);
            if (!m.Success)
            {
                throw new ArgumentException("The given specifier does not denote a Literal");
            }
            this.specifier = specifier;
        }

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
                if (specifier.Length > 1)
                {
                    string relevant = specifier.Substring(1); // skip the damn Backslash at the start
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
                        // special escapes here
                        switch (relevant[0])
                        {
                            case 'd':
                                return AssistantResources.AtomDescription_Digit;
                            case 'D':
                                return AssistantResources.AtomDescription_NonDigit;
                            case 'b':
                                return AssistantResources.AtomDescription_WordBoundary;
                            case 'B':
                                return AssistantResources.AtomDescription_NonWordBoundary;
                            case 'w':
                                return AssistantResources.AtomDescription_WordCharacter;
                            case 'W':
                                return AssistantResources.AtomDescription_NonWordCharacter;
                            case 's':
                                return AssistantResources.AtomDescription_Whitespace;
                            case 'S':
                                return AssistantResources.AtomDescription_NonWhitespace;
                            case 'n':
                                return AssistantResources.AtomDescription_Newline;
                            case 'r':
                                return AssistantResources.AtomDescription_CarriageReturn;
                            case 'f':
                                return AssistantResources.AtomDescription_FormFeed;
                            case 'v':
                                return AssistantResources.AtomDescription_VTab;
                            case 't':
                                return AssistantResources.AtomDescription_HTab;
                            default:
                                // shouldn't ever happen, so we blow it all up
                                throw new InvalidOperationException("took an escape sequence that shouldn't exist");
                        }
                    }
                }
                else
                {
                    if (specifier.Equals("."))
                    {
                        return AssistantResources.AtomDescription_Dot;
                    }
                    // Behaviour with "." needs fix
                    return string.Format(AssistantResources.AtomDescription_Literal_ActualLiteral, specifier);
                }

                throw new NotImplementedException();
            }
        }

        public bool TryMatch(ref string text)
        {
            throw new NotImplementedException();
        }
    }
}
