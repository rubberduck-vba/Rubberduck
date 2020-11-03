using System;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField.Extensions
{
    public static class StringExtensions
    {
        public static bool IsEquivalentVBAIdentifierTo(this string lhs, string identifier)
            => lhs.Equals(identifier, StringComparison.InvariantCultureIgnoreCase);

        public static string IncrementEncapsulationIdentifier(this string identifier)
        {
            var numeric = string.Concat(identifier.Reverse().TakeWhile(c => char.IsDigit(c)).Reverse());
            if (!int.TryParse(numeric, out var currentNum))
            {
                currentNum = 0;
            }
            var identifierSansNumericSuffix = identifier.Substring(0, identifier.Length - numeric.Length);
            return $"{identifierSansNumericSuffix}{++currentNum}";
        }
    }
}
