using Rubberduck.RegexAssistant.i18n;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Rubberduck.RegexAssistant.Atoms
{
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
