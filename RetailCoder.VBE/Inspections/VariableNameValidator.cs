using System.Globalization;
using Rubberduck.Parsing.Grammar;
using System;
using System.Linq;

namespace Rubberduck.Inspections
{
    public class VariableNameValidator
    {
        public VariableNameValidator() { }
        public VariableNameValidator(string identifier) { _identifier = identifier; }

        private const string AllVowels = "aeiouyàâäéèêëïîöôùûü";
        private const int MinVariableNameLength = 3;

        /****  Meaningful Name Characteristics  ************/
        public bool HasVowels
        {
            get
            {
                return _identifier.Any(character => AllVowels.Any(vowel =>
                    string.Compare(vowel.ToString(CultureInfo.InvariantCulture),
                        character.ToString(CultureInfo.InvariantCulture), StringComparison.OrdinalIgnoreCase) == 0));
            }
        }

        public bool HasConsonants
        {
            get
            {
                return !_identifier.All(character => AllVowels.Any(vowel =>
                    string.Compare(vowel.ToString(CultureInfo.InvariantCulture),
                        character.ToString(CultureInfo.InvariantCulture), StringComparison.OrdinalIgnoreCase) == 0));
            }
        }

        public bool IsSingleRepeatedLetter
        {
            get
            {
                var firstLetter = _identifier.First().ToString(CultureInfo.InvariantCulture);
                return _identifier.All(a => string.Compare(a.ToString(CultureInfo.InvariantCulture), firstLetter,
                    StringComparison.OrdinalIgnoreCase) == 0);
            }
        }

        public bool IsTooShort { get { return _identifier.Length < MinVariableNameLength; } }
        public bool EndsWithNumber { get { return char.IsDigit(_identifier.Last()); } }

        /****  Invalid Name Characteristics  ************/
        public bool StartsWithNumber { get { return FirstLetterIsDigit(); } }

        public bool IsReservedName
        {
            get
            {
                var tokenValues = typeof(Tokens).GetFields().Select(item => item.GetValue(null)).Cast<string>().Select(item => item);
                return tokenValues.Contains(_identifier, StringComparer.InvariantCultureIgnoreCase);
            }
        }

        public bool ContainsSpecialCharacters { get { return UsesSpecialCharacters(); } }

        private string _identifier;
        public string Identifier
        {
            get { return _identifier; }
            set { _identifier = value; }
        } 

        public bool IsValidName()
        {
            return !string.IsNullOrEmpty(_identifier) 
                && !StartsWithNumber
                && !IsReservedName
                && !ContainsSpecialCharacters;
        }

        public bool IsMeaningfulName()
        {
            return HasVowels
                && HasConsonants
                && !IsSingleRepeatedLetter
                && !IsTooShort
                && !EndsWithNumber;
        }

        private bool FirstLetterIsDigit()
        {
            return !char.IsLetter(_identifier.FirstOrDefault());
        }

        private bool UsesSpecialCharacters()
        {
            return _identifier.Any(c => !char.IsLetterOrDigit(c) && c != '_');
        }
    }
}
