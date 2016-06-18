using Rubberduck.RegexAssistant.Extensions;
using Rubberduck.RegexAssistant.i18n;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Rubberduck.RegexAssistant
{
    public interface Atom : IRegularExpression
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
                throw new ArgumentException("The give specifier does not denote a character class");
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
                return string.Format(InverseMatching ? AssistantResources.AtomDescription_CharacterClass_Inverted : AssistantResources.AtomDescription_CharacterClass, HumanReadableClass(), Quantifier.HumanReadable());
            }
        }

        private string HumanReadableClass()
        {
            return string.Join(", ", CharacterSpecifiers); // join last with and?
        }

        public Quantifier Quantifier
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public bool TryMatch(ref string text)
        {
            throw new NotImplementedException();
        }
    }

    class Group : Atom
    {
        public string Description
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public Quantifier Quantifier
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public bool TryMatch(ref string text)
        {
            throw new NotImplementedException();
        }
    }

    class EscapedCharacter : Atom
    {
        public string Description
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public Quantifier Quantifier
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public bool TryMatch(ref string text)
        {
            throw new NotImplementedException();
        }
    }

    class Literal : Atom
    {
        public string Description
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public Quantifier Quantifier
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public bool TryMatch(ref string text)
        {
            throw new NotImplementedException();
        }
    }
}
