using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public static class StringExtensions
    {
        public static string Capitalize(this string input)
            => $"{char.ToUpperInvariant(input[0]) + input.Substring(1, input.Length - 1)}";

        public static string UnCapitalize(this string input)
            => $"{char.ToLowerInvariant(input[0]) + input.Substring(1, input.Length - 1)}";

        public static bool EqualsVBAIdentifier(this string lhs, string identifier)
            => lhs.Equals(identifier, StringComparison.InvariantCultureIgnoreCase);
    }
}
