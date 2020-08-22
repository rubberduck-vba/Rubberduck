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
            var fragments = identifier.Split('_');
            if (fragments.Length == 1) { return $"{identifier}_1"; }

            var lastFragment = fragments[fragments.Length - 1];
            if (long.TryParse(lastFragment, out var number))
            {
                fragments[fragments.Length - 1] = (number + 1).ToString();

                return string.Join("_", fragments);
            }
            return $"{identifier}_1"; ;
        }

        public static string LimitNewlines(this string content, int maxConsecutiveNewlines = 2)
        {
            var target = string.Concat(Enumerable.Repeat(Environment.NewLine, maxConsecutiveNewlines + 1).ToList());
            var replacement = string.Concat(Enumerable.Repeat(Environment.NewLine, maxConsecutiveNewlines).ToList());
            var guard = 0;
            var maxAttempts = 100;
            while (++guard < maxAttempts && content.Contains(target))
            {
                content = content.Replace(target, replacement);
            }

            if (guard >= maxAttempts)
            {
                throw new FormatException($"Unable to limit consecutive '{Environment.NewLine}' strings to {maxConsecutiveNewlines}");
            }
            return content;
        }
    }
}
