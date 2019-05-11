using System;
using System.Globalization;

namespace Rubberduck.Common
{
    public static class StringExtensions
    {
        public static string Capitalize(this string input)
        {
            var tokens = input.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length == 0)
            {
                return input;
            }
            tokens[0] = CultureInfo.CurrentUICulture.TextInfo.ToTitleCase(tokens[0]);
            return string.Join(" ", tokens);
        }
        public static string CapitalizeFirstLetter(this string input)
         {
            if (input.Length == 0)
            {
                return string.Empty;
            }
            return input.Capitalize().Substring(0, 1) + input.Substring(1);
         }
    }
}
