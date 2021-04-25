using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using Tokens = Rubberduck.Resources.Tokens;

namespace Rubberduck.Refactorings.Common
{
    public static class VBAIdentifierValidator
    {

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
            const string Vowels = "aeiouyàâäéèêëïîöôùûü";
            const int MinimumNameLength = 3;

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
        public static bool IsValidIdentifier(string name, DeclarationType declarationType, bool isArrayDeclaration = false)
            => !TryMatchInvalidIdentifierCriteria(name, declarationType, out _, isArrayDeclaration);

        /// <summary>
        /// Evaluates an identifier string's conformance with MS-VBAL naming requirements.
        /// </summary>
        /// <returns>Message for first matching invalid identifier criteria.  Or, an empty string if the identifier is valid</returns>
        public static bool TryMatchInvalidIdentifierCriteria(string name, DeclarationType declarationType, out string criteriaMatchMessage, bool isArrayDeclaration = false)
        {
            criteriaMatchMessage = string.Empty;

            var maxNameLength = declarationType.HasFlag(DeclarationType.Module)
               ? Declaration.MaxModuleNameLength : Declaration.MaxMemberNameLength;

            if (string.IsNullOrEmpty(name))
            {
                criteriaMatchMessage = RefactoringsUI.InvalidNameCriteria_IsNullOrEmpty;
                return true;
            }

            //Does not start with a letter
            if (!char.IsLetter(name.First()))
            {
                criteriaMatchMessage = string.Format(RefactoringsUI.InvalidNameCriteria_DoesNotStartWithLetterFormat, name);
                return true;
            }

            //Has special characters
            if (name.Any(c => !char.IsLetterOrDigit(c) && c != '_'))
            {
                criteriaMatchMessage = string.Format(RefactoringsUI.InvalidNameCriteria_InvalidCharactersFormat, name);
                return true;
            }

            //Is a reserved identifier
            if (!declarationType.HasFlag(DeclarationType.UserDefinedTypeMember))
            {
                if (Tokens.IllegalIdentifierNames.Contains(name, StringComparer.InvariantCultureIgnoreCase))
                {
                    criteriaMatchMessage = string.Format(RefactoringsUI.InvalidNameCriteria_IsReservedKeywordFormat, name);
                    return true;
                }
            }
            else if (isArrayDeclaration) //is a DeclarationType.UserDefinedTypeMember
            {
                //DeclarationType.UserDefinedTypeMember can have reserved identifier keywords
                //...unless the declaration is an array.  Adding the parentheses causes errors.

                //Name is not a reserved identifier, but when used as a UDTMember array declaration
                //it collides with the 'Name' Statement (Renames a disk file, directory, or folder)
                var invalidUDTArrayIdentifiers = Tokens.IllegalIdentifierNames.Concat(new List<string>() { "Name" });

                if (invalidUDTArrayIdentifiers.Contains(name, StringComparer.InvariantCultureIgnoreCase))
                {
                    criteriaMatchMessage = string.Format(RefactoringsUI.InvalidNameCriteria_IsReservedKeywordFormat, name);
                    return true;
                }
            }

            //"VBA" identifier not allowed for projects
            if (declarationType.HasFlag(DeclarationType.Project)
                && name.Equals("VBA", StringComparison.InvariantCultureIgnoreCase))
            {
                criteriaMatchMessage = string.Format(RefactoringsUI.InvalidNameCriteria_IsReservedKeywordFormat, name);
                return true;
            }

            //Exceeds max length
            if (name.Length > maxNameLength)
            {
                criteriaMatchMessage = string.Format(RefactoringsUI.InvalidNameCriteria_ExceedsMaximumLengthFormat, name);
                return true;
            }
            return false;
        }

        /// <summary>
        /// Evaluates an identifier string's conformance with MS-VBAL naming requirements.
        /// </summary>
        /// <returns>Messages for all matching invalid identifier criteria</returns>
        public static IReadOnlyList<string> SatisfiedInvalidIdentifierCriteria(string name, DeclarationType declarationType, bool isArrayDeclaration = false)
        {
            var criteriaMatchMessages = new List<string>();

            var maxNameLength = declarationType.HasFlag(DeclarationType.Module)
               ? Declaration.MaxModuleNameLength 
               : Declaration.MaxMemberNameLength;

            if (string.IsNullOrEmpty(name))
            {
                criteriaMatchMessages.Add(RefactoringsUI.InvalidNameCriteria_IsNullOrEmpty);
            }

            //Does not start with a letter
            if (!char.IsLetter(name.First()))
            {
                criteriaMatchMessages.Add(string.Format(RefactoringsUI.InvalidNameCriteria_DoesNotStartWithLetterFormat, name));
            }

            //Has special characters
            if (name.Any(c => !char.IsLetterOrDigit(c) && c != '_'))
            {
                criteriaMatchMessages.Add(string.Format(RefactoringsUI.InvalidNameCriteria_InvalidCharactersFormat, name));
            }

            //Is a reserved identifier
            if (!declarationType.HasFlag(DeclarationType.UserDefinedTypeMember))
            {
                if (Tokens.IllegalIdentifierNames.Contains(name, StringComparer.InvariantCultureIgnoreCase))
                {
                    criteriaMatchMessages.Add(string.Format(RefactoringsUI.InvalidNameCriteria_IsReservedKeywordFormat, name));
                }
            }
            else if (isArrayDeclaration) //is a DeclarationType.UserDefinedTypeMember
            {
                //DeclarationType.UserDefinedTypeMember can have reserved identifier keywords
                //...unless the declaration is an array.  Adding the parentheses causes errors.

                //Name is not a reserved identifier, but when used as a UDTMember array declaration
                //it collides with the 'Name' Statement (Renames a disk file, directory, or folder)
                var invalidUDTArrayIdentifiers = Tokens.IllegalIdentifierNames.Concat(new List<string>() { "Name" });

                if (invalidUDTArrayIdentifiers.Contains(name, StringComparer.InvariantCultureIgnoreCase))
                {
                    criteriaMatchMessages.Add(string.Format(RefactoringsUI.InvalidNameCriteria_IsReservedKeywordFormat, name));
                }
            }

            //"VBA" identifier not allowed for projects
            if (declarationType.HasFlag(DeclarationType.Project)
                && name.Equals("VBA", StringComparison.InvariantCultureIgnoreCase))
            {
                criteriaMatchMessages.Add(string.Format(RefactoringsUI.InvalidNameCriteria_IsReservedKeywordFormat, name));
            }

            //Exceeds max length
            if (name.Length > maxNameLength)
            {
                criteriaMatchMessages.Add(string.Format(RefactoringsUI.InvalidNameCriteria_ExceedsMaximumLengthFormat, name));
            }

            return criteriaMatchMessages;
        }
    }
}
