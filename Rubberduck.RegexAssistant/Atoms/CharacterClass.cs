using Rubberduck.RegexAssistant.i18n;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Rubberduck.Resources;

namespace Rubberduck.RegexAssistant.Atoms
{
    internal class CharacterClass : IAtom
    {
        public bool InverseMatching { get; }
        public IList<string> CharacterSpecifiers { get; }

        public CharacterClass(string specifier, Quantifier quantifier, bool spellOutWhiteSpace = false)
        {
            if (specifier == null || quantifier == null)
            {
                throw new ArgumentNullException();
            }

            Quantifier = quantifier;
            if (!specifier.StartsWith("[") || !specifier.EndsWith("]"))
            {
                throw new ArgumentException("The given specifier does not denote a character class");
            }
            Specifier = specifier;
            // trim leading and closing bracket
            var actualSpecifier = specifier.Substring(1, specifier.Length - 2);
            InverseMatching = actualSpecifier.StartsWith("^");
            CharacterSpecifiers = ExtractCharacterSpecifiers(InverseMatching 
                    ? actualSpecifier.Substring(1) 
                    : actualSpecifier
                , spellOutWhiteSpace);
        }

        public string Specifier { get; }

        public Quantifier Quantifier { get; }

        private static readonly Regex CharacterRanges = new Regex(@"(\\[dDwWsS]|(\\[ntfvr]|\\([0-7]{3}|x[\dA-F]{2}|u[\dA-F]{4}|[\\\.\[\]])|.)(-(\\[ntfvr]|\\([0-7]{3}|x[A-F]{2}|u[\dA-F]{4}|[\.\\\[\]])|.))?)", RegexOptions.Compiled);
        
        private IList<string> ExtractCharacterSpecifiers(string characterClass, bool spellOutWhitespace)
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

                result.Add(spellOutWhitespace && WhitespaceToString.IsFullySpellingOutApplicable(specifier.Value, out var spelledOutWhiteSpace)
                    ? spelledOutWhiteSpace
                    : specifier.Value);
            }
            return result;
        }

        public string Description(bool spellOutWhitespace) => string.Format(InverseMatching 
                ? AssistantResources.AtomDescription_CharacterClass_Inverted 
                : AssistantResources.AtomDescription_CharacterClass
            , HumanReadableClass());

        private string HumanReadableClass()
        {
            return string.Join(", ", CharacterSpecifiers); // join last with and?
        }

        public override string ToString() => Specifier;
        public override bool Equals(object obj)
        {
            return obj is CharacterClass other 
                && other.Quantifier.Equals(Quantifier)
                && other.Specifier.Equals(Specifier);
        }
        public override int GetHashCode() => HashCode.Compute(Specifier, Quantifier);
    }
}
