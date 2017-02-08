using Rubberduck.Parsing.Grammar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Inspections
{
    public class VariableNameValidator
    {
        public VariableNameValidator() { }
        public VariableNameValidator(string identifier) { _identifier = identifier; }


        private const string _ALL_VOWELS = "aeiouyàâäéèêëïîöôùûü";
        private const int _MIN_VARIABLE_NAME_LENGTH = 3;

        /****  Meaningful Name Characteristics  ************/
        public bool HasVowels { get { return hasVowels(); } }
        public bool HasConsonants { get { return hasConsonants(); } }
        public bool IsSingleRepeatedLetter { get { return nameIsASingleRepeatedLetter(); } }
        public bool IsTooShort { get { return _identifier.Length < _MIN_VARIABLE_NAME_LENGTH; } }
        public bool EndsWithNumber { get { return endsWithNumber(); } }

        /****  Invalid Name Characteristics  ************/
        public bool StartsWithNumber { get { return FirstLetterIsDigit(); } }
        public bool IsReservedName { get { return isReservedName(); } }
        public bool ContainsSpecialCharacters { get { return UsesSpecialCharacters(); } }

        private string _identifier;
        public string Identifier
        {
            get { return _identifier; }
            set { _identifier = value; }
        } 
        public bool IsValidName()
        {
            return !StartsWithNumber
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
        private bool IsEmpty()
        {
            return _identifier.Equals(string.Empty);
        }
        private bool FirstLetterIsDigit()
        {
            return !char.IsLetter(_identifier.FirstOrDefault());
       }
        private bool isReservedName()
        {
            var tokenValues = typeof(Tokens).GetFields().Select(item => item.GetValue(null)).Cast<string>().Select(item => item);
            return tokenValues.Contains(_identifier, StringComparer.InvariantCultureIgnoreCase);
        }
        private bool UsesSpecialCharacters()
        {
            return _identifier.Any(c => !char.IsLetterOrDigit(c) && c != '_');
        }

        private bool hasVowels()
        {
            return _identifier.Any(character => _ALL_VOWELS.Any(vowel =>
                   string.Compare(vowel.ToString(), character.ToString(), StringComparison.OrdinalIgnoreCase) == 0));
        }
        private bool hasConsonants()
        {
            return !_identifier.All(character => _ALL_VOWELS.Any(vowel =>
                   string.Compare(vowel.ToString(), character.ToString(), StringComparison.OrdinalIgnoreCase) == 0));
        }
        private bool nameIsASingleRepeatedLetter()
        {
            string firstLetter = _identifier.First().ToString();
            return _identifier.All(a => string.Compare(a.ToString(), firstLetter,
                StringComparison.OrdinalIgnoreCase) == 0);
        }
        private bool endsWithNumber()
        {
            return char.IsDigit(_identifier.Last());
        }
    }
}
