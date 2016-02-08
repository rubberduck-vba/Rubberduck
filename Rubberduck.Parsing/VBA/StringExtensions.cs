using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.VBA
{
    public static class StringExtensions
    {


        /// <summary>
        /// Returns a value indicating whether line of code is/contains a comment.
        /// </summary>
        /// <param name="line">The extended string.</param>
        /// <param name="index">The start index of the comment string.</param>
        /// <returns>Returns <c>true</c> if specified string contains a VBA comment marker outside a string literal.</returns>
        public static bool HasComment(this string line, out int index)
        {
            var instruction = line.StripStringLiterals();

            index = instruction.IndexOf(Tokens.CommentMarker, StringComparison.InvariantCulture);
            if (index >= 0)
            {
                return true;
            }

            index = instruction.IndexOf(Tokens.Rem + " ", StringComparison.InvariantCulture);
            return index >= 0;
        }

        public static string StripStringLiterals(this string line)
        {
            return Regex.Replace(line, "\"[^\"]*\"", match => new string(' ', match.Length));
        }

        public static string RemoveExtraSpacesLeavingIndentation(this string line)
        {
            var newString = new StringBuilder();
            var lastWasWhiteSpace = false;

            if (line.All(char.IsWhiteSpace))
            {
                return line;
            }

            var firstNonwhitespaceIndex = line.IndexOf(line.FirstOrDefault(c => !char.IsWhiteSpace(c)));

            for (var i = 0; i < line.Length; i++)
            {
                if (i < firstNonwhitespaceIndex)
                {
                    newString.Append(line[i]);
                    continue;
                }

                if (char.IsWhiteSpace(line[i]) && lastWasWhiteSpace) { continue; }

                newString.Append(line[i]);
                lastWasWhiteSpace = char.IsWhiteSpace(line[i]);
            }

            return newString.ToString().Replace('\r', ' ');
        }

        public static int NthIndexOf(this string line, char chr, int index)
        {
            var currentIndexOf = 0;

            for (var i = 0; i < line.Length; i++)
            {
                if (line[i] == chr)
                {
                    currentIndexOf++;
                }

                if (currentIndexOf == index)
                {
                    return i;
                }
            }

            throw new ArgumentException(string.Format("Not {0} instances of '{1}' in '{2}'", index, chr, line), "index");
        }
    }
}
