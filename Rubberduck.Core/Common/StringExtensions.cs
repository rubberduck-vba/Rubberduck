using Rubberduck.VBEditor;
using System;
using System.Globalization;

namespace Rubberduck.Common
{
    public static class StringExtensions
    {
        public static CodeString ToCodeString(this string code)
        {
            var zPosition = new Selection();
            var lines = (code ?? string.Empty).Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                var line = lines[i];
                var index = line.IndexOf('|');
                if (index >= 0)
                {
                    lines[i] = line.Remove(index, 1);
                    zPosition = new Selection(i, index);
                    break;
                }
            }

            var newCode = string.Join("\n", lines);
            return new CodeString(newCode, zPosition);
        }

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
