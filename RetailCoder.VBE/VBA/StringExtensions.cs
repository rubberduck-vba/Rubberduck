using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace Rubberduck.VBA.Grammar
{
    [ComVisible(false)]
    public static class StringExtensions
    {
        public static readonly char StringDelimiter = '"';
        public static readonly char CommentMarker = '\'';

        /// <summary>
        /// Strips any trailing comment from specified line of code.
        /// </summary>
        /// <param name="line"></param>
        /// <returns>Returns a new string, without the trailing comment.</returns>
        public static string StripTrailingComment(this string line)
        {
            int index;
            if (line.HasComment(out index))
            {
                if (index == 0)
                {
                    return string.Empty;// line;
                }

                return line.EndsWith(":") 
                    ? line.Substring(0, index - 2) 
                    : line.Substring(0, index).TrimEnd();
            }

            return line;
        }

        /// <summary>
        /// Returns a value indicating whether line of code is/contains a comment.
        /// </summary>
        /// <param name="line">The extended string.</param>
        /// <param name="index">The start index of the comment string.</param>
        /// <returns>Returns <c>true</c> if specified string contains a VBA comment marker.</returns>
        public static bool HasComment(this string line, out int index)
        {
            var instruction = line.StripStringLiterals();

            index = instruction.IndexOf(CommentMarker);
            if (index >= 0)
            {
                return true;
            }

            index = instruction.IndexOf(ReservedKeywords.Rem + " ", StringComparison.InvariantCulture);
            return index >= 0;
        }

        public static string StripStringLiterals(this string line)
        {
            return Regex.Replace(line, "\"[^\"]*\"", match => new string(' ', match.Length));
        }

        /// <summary>
        /// Strips all string literals from a line of code or instruction.
        /// Replaces string literals with whitespace characters, to maintain original length.
        /// </summary>
        /// <param name="line"></param>
        /// <returns>Returns a new string, stripped of all string literals and string delimiters.</returns>
        //public static string StripStringLiterals(this string line)
        //{
        //    var builder = new StringBuilder(line.Length);
        //    var isInsideString = false;
        //    for (var cursor = 0; cursor < line.Length; cursor++)
        //    {
        //        if (line[cursor] == StringDelimiter)
        //        {
        //            if (isInsideString)
        //            {
        //                isInsideString = cursor + 1 == line.Length || line[cursor + 1] == StringDelimiter || cursor > 0 && (line[cursor - 1] == StringDelimiter);
        //            }
        //            else
        //            {
        //                isInsideString = true;
        //            }
        //        }

        //        if (!isInsideString && line[cursor] != StringDelimiter)
        //        {
        //            builder.Append(line[cursor]);
        //        }
        //        else
        //        {
        //            builder.Append(' ');
        //        }
        //    }

        //    return builder.ToString();
        //}
    }
}
