﻿using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Rubberduck.Refactorings.Common
{
    public static class VBAIdentifierValidator
    {
        private static IEnumerable<string> ReservedIdentifiers =
            typeof(Tokens).GetFields().Select(item => item.GetValue(null).ToString()).ToArray();

        /// <summary>
        /// Predicate function determining if an identifier string's content will trigger a result for the UseMeaningfulNames inspection.
        /// </summary>
        public static bool IsMeaningfulIdentifier(string name)
                => !TryMatchMeaninglessIdentifierCriteria(name, out _);

        /// <summary>
        /// Evaluates if an identifier string's content will trigger a result for the UseMeaningfulNames inspection.
        /// </summary>
        /// <returns>Message indicating that the string will result in a UseMeaningfulNames inspection result</returns>
        public static bool TryMatchMeaninglessIdentifierCriteria(string name, out string criteriaMatchMessage)
        {
            criteriaMatchMessage = string.Empty;
            string Vowels = "aeiouyàâäéèêëïîöôùûü";
            int MinimumNameLength = 3;

            bool HasVowel()
                => name.Any(character => Vowels.Any(vowel =>
                    string.Compare(vowel.ToString(CultureInfo.InvariantCulture),
                        character.ToString(CultureInfo.InvariantCulture), StringComparison.OrdinalIgnoreCase) == 0));

            bool HasConsonant()
                => !name.All(character => Vowels.Any(vowel =>
                    string.Compare(vowel.ToString(CultureInfo.InvariantCulture),
                        character.ToString(CultureInfo.InvariantCulture), StringComparison.OrdinalIgnoreCase) == 0));

            bool IsRepeatedCharacter()
            {
                var firstLetter = name.First().ToString(CultureInfo.InvariantCulture);
                return name.All(a => string.Compare(a.ToString(CultureInfo.InvariantCulture), firstLetter,
                    StringComparison.OrdinalIgnoreCase) == 0);
            }

            bool EndsWithDigit()
                => char.IsDigit(name.Last());

            bool IsTooShort()
                => name.Length < MinimumNameLength;

            var isMeaningless = !(HasVowel()
                                    && HasConsonant()
                                    && !IsRepeatedCharacter()
                                    && !IsTooShort()
                                    && !EndsWithDigit());

            if (isMeaningless)
            {
                criteriaMatchMessage = string.Format(RubberduckUI.MeaninglessNameCriteriaMatchFormat, name);
            }
            return isMeaningless;
        }

        /// <summary>
        /// Predicate function determining if an identifier string conforms to MS-VBAL naming requirements
        /// </summary>
        public static bool IsValidIdentifier(string name, DeclarationType declarationType)
            => !TryMatchInvalidIdentifierCriteria(name, declarationType, out _);

        /// <summary>
        /// Evaluates an identifier string's conformance with MS-VBAL naming requirements.
        /// </summary>
        /// <returns>Message for first matching invalid identifier criteria.  Or, an empty string if the identifier is valid</returns>
        public static bool TryMatchInvalidIdentifierCriteria(string name, DeclarationType declarationType, out string criteriaMatchMessage)
        {
            criteriaMatchMessage = string.Empty;

            var maxNameLength = declarationType.HasFlag(DeclarationType.Module)
               ? Declaration.MaxModuleNameLength : Declaration.MaxMemberNameLength;

            if (string.IsNullOrEmpty(name))
            {
                criteriaMatchMessage = RubberduckUI.InvalidNameCriteria_IsNullOrEmpty;
                return true;
            }

            //Does not start with a letter
            if (!char.IsLetter(name.First()))
            {
                criteriaMatchMessage = string.Format(RubberduckUI.InvalidNameCriteria_DoesNotStartWithLetterFormat, name);
                return true;
            }

            //Has special characters
            if (name.Any(c => !char.IsLetterOrDigit(c) && c != '_'))
            {
                criteriaMatchMessage = string.Format(RubberduckUI.InvalidNameCriteria_InvalidCharactersFormat, name);
                return true;
            }

            //Is a reserved identifier
            if (ReservedIdentifiers.Contains(name, StringComparer.InvariantCultureIgnoreCase))
            {
                criteriaMatchMessage = string.Format(RubberduckUI.InvalidNameCriteria_IsReservedKeywordFormat, name);
                return true;
            }

            //"VBA" identifier not allowed for projects
            if (declarationType.HasFlag(DeclarationType.Project)
                && name.Equals("VBA", StringComparison.InvariantCultureIgnoreCase))
            {
                criteriaMatchMessage = string.Format(RubberduckUI.InvalidNameCriteria_IsReservedKeywordFormat, name);
                return true;
            }

            //Exceeds max length
            if (name.Length > maxNameLength)
            {
                criteriaMatchMessage = string.Format(RubberduckUI.InvalidNameCriteria_ExceedsMaximumLengthFormat, name);
                return true;
            }
            return false;
        }
    }
}
