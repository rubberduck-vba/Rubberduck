using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Common
{
    public static class VariableNameValidator
    {
        private static readonly string Vowels = "aeiouyàâäéèêëïîöôùûü";
        private static readonly int MinVariableNameLength = 3;

        private static bool HasVowel(string name)
        {
            return name.Any(character => Vowels.Any(vowel =>
                string.Compare(vowel.ToString(CultureInfo.InvariantCulture),
                    character.ToString(CultureInfo.InvariantCulture), StringComparison.OrdinalIgnoreCase) == 0));
        }

        private static bool HasConsonant(string name)
        {
            return !name.All(character => Vowels.Any(vowel =>
                string.Compare(vowel.ToString(CultureInfo.InvariantCulture),
                    character.ToString(CultureInfo.InvariantCulture), StringComparison.OrdinalIgnoreCase) == 0));
        }

        private static bool IsRepeatedCharacter(string name)
        {
            var firstLetter = name.First().ToString(CultureInfo.InvariantCulture);
            return name.All(a => string.Compare(a.ToString(CultureInfo.InvariantCulture), firstLetter,
                StringComparison.OrdinalIgnoreCase) == 0);
        }

        private static bool IsUnderMinLength(string name)
        {
            return name.Length < MinVariableNameLength;
        }

        private static bool EndsWithDigit(string name)
        {
            return char.IsDigit(name.Last());
        }

        public static bool StartsWithDigit(string name)
        {
            return !char.IsLetter(name.First());
        }

        private static readonly IEnumerable<string> ReservedNames =
            typeof (Tokens).GetFields().Select(item => item.GetValue(null).ToString()).ToArray();

        public static bool IsReservedIdentifier(string name)
        {
            return ReservedNames.Contains(name, StringComparer.InvariantCultureIgnoreCase);
        }

        public static bool HasSpecialCharacters(string name)
        {
            return name.Any(c => !char.IsLetterOrDigit(c) && c != '_');
        }

        public static bool IsValidName(string name)
        {
            return !string.IsNullOrEmpty(name)
                   && !StartsWithDigit(name)
                   && !IsReservedIdentifier(name)
                   && !HasSpecialCharacters(name);
        }

        public static bool IsMeaningfulName(string name)
        {
            return HasVowel(name)
                && HasConsonant(name)
                && !IsRepeatedCharacter(name)
                && !IsUnderMinLength(name)
                && !EndsWithDigit(name);
        }
    }
}
