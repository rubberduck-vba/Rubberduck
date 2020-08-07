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

        public static string LimitNewLines(this string content, int maxConsecutiveNewLines)
        {
            if (maxConsecutiveNewLines < 2) { throw new ArgumentOutOfRangeException(); }

            var target = string.Concat(Enumerable.Repeat(Environment.NewLine, maxConsecutiveNewLines + 1).ToList());
            var replacement = string.Concat(Enumerable.Repeat(Environment.NewLine, maxConsecutiveNewLines).ToList());

            for (var counter = 1; counter < 100 && content.Contains(target); counter++)
            {
                content = content.Replace(target, replacement);
            }
            return content;
        }
    }
}
