using System;
using System.Text.RegularExpressions;

namespace Rubberduck.VBA.Grammar
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
    }
}
