using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Common
{
    public static class StringExtensions
    {
        public static string Captialize(this string input)
        {
            var tokens = input.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length == 0)
            {
                return input;
            }
            tokens[0] = CultureInfo.CurrentUICulture.TextInfo.ToTitleCase(tokens[0]);
            return string.Join(" ", tokens);
        }
    }
}
