using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

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

        public static string ToLowerCase(this string input)
        {
            var tokens = input.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length == 0)
            {
                return input;
            }
            tokens[0] = CultureInfo.CurrentUICulture.TextInfo.ToLower(tokens[0]);
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

        public static string ToLowerCaseFirstLetter(this string input)
        {
            if (input.Length == 0)
            {
                return string.Empty;
            }
            return input.ToLowerCase().Substring(0, 1) + input.Substring(1);
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

        public static string FromVbaStringLiteral(this string input)
        {
            return input.UnQuote().Replace("\"\"", "\"");
        }

        public static string ToVbaStringLiteral(this string input)
        {
            return $"\"{input.Replace("\"", "\"\"")}\"";
        }

        public static bool TryMatchHungarianNotationCriteria(this string identifier, out string nonHungarianName)
        {
            nonHungarianName = identifier;
            if (HungarianIdentifierRegex.IsMatch(identifier))
            {
                var prefixChars = identifier.TakeWhile(c => char.IsLower(c));
                nonHungarianName = identifier.Substring(prefixChars.Count());
                return true;
            }
            return false;
        }

        private static readonly List<string> HungarianPrefixes = new List<string>
        {
            "chk",
            "cbo",
            "cmd",
            "btn",
            "fra",
            "img",
            "lbl",
            "lst",
            "mnu",
            "opt",
            "pic",
            "shp",
            "txt",
            "tmr",
            "chk",
            "dlg",
            "drv",
            "frm",
            "grd",
            "obj",
            "rpt",
            "fld",
            "idx",
            "tbl",
            "tbd",
            "bas",
            "cls",
            "g",
            "m",
            "bln",
            "byt",
            "col",
            "dtm",
            "dbl",
            "cur",
            "int",
            "lng",
            "sng",
            "str",
            "udt",
            "vnt",
            "var",
            "pgr",
            "dao",
            "b",
            "by",
            "c",
            "chr",
            "i",
            "l",
            "s",
            "o",
            "n",
            "dt",
            "dat",
            "a",
            "arr"
        };

        private static readonly Regex HungarianIdentifierRegex = new Regex($"^({string.Join("|", HungarianPrefixes)})[A-Z0-9].*$");
    }
}
