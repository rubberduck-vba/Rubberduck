using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.VBA
{
    public static class StringExtensions
    {
        // see issues #1057 and #2364.
        private static readonly IList<string> ValidRemCommentMarkers =
            new List<string>
            {
                Tokens.Rem + ' ',
                Tokens.Rem + '?',
                Tokens.Rem + '<',
                Tokens.Rem + '>',
                Tokens.Rem + '{',
                Tokens.Rem + '}',
                Tokens.Rem + '~',
                Tokens.Rem + '`',
                Tokens.Rem + '!',
                Tokens.Rem + '/',
                Tokens.Rem + '*',
                Tokens.Rem + '(',
                Tokens.Rem + ')',
                Tokens.Rem + '-',
                Tokens.Rem + '=',
                Tokens.Rem + '+',
                Tokens.Rem + '\\',
                Tokens.Rem + '|',
                Tokens.Rem + ';',
                Tokens.Rem + ':',
                Tokens.Rem + '\'',
                Tokens.Rem + '"',
                Tokens.Rem + ',',
                Tokens.Rem + '.',
            };

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
                // line contains a single-quote comment marker
                return true;
            }

            // note: REM comment markers are NOT implemented as per language specifications.
            // ReSharper disable once ForCanBeConvertedToForeach
            for (var i = 0; i < ValidRemCommentMarkers.Count; i++)
            {
                index = instruction.IndexOf(ValidRemCommentMarkers[i], StringComparison.InvariantCulture);
                if (index >= 0)
                {
                    return true;
                }
            }

            return false;
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

            throw new ArgumentException(string.Format(ParsingText.StringExtensionsArgumentException, index, chr, line), "index");
        }
    }
}
