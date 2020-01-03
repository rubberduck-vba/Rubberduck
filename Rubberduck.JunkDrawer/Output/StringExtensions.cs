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

        public static string UnQuote(this string input)
        {
            if (input[0] == '"' && input[input.Length - 1] == '"')
            {
                return input.Substring(1, input.Length - 2);
            }
            return input;
        }

        public static string EnQuote(this string input)
        {
            if (input[0] == '"' && input[input.Length - 1] == '"')
            {
                return input;
            }
            return $"\"{input}\"";
        }
    }
}
