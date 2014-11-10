using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RetailCoderVBE.Reflection.VBA.Grammar
{
    internal static class StringExtensions
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
                return line.Substring(0, index - 1).TrimEnd();
            }

            return line;
        }

        /// <summary>
        /// Returns a value indicating whether line of code is/contains a comment.
        /// </summary>
        /// <param name="line"></param>
        /// <param name="comment">Returns the comment string, including the comment marker.</param>
        /// <returns></returns>
        public static bool HasComment(this string line, out int index)
        {
            var result = false;

            var isString = false;
            index = -1;

            for (int cursor = 0; cursor < line.Length - 1; cursor++)
            {
                // determine if cursor is inside a string literal:
                if (line[cursor] == StringDelimiter)
                {
                    if (isString)
                    {
                        isString = line[cursor + 1] == StringDelimiter || cursor > 0 && (line[cursor - 1] == StringDelimiter);
                    }
                    else
                    {
                        isString = true;
                    }
                }

                if (!isString && line[cursor] == CommentMarker)
                {
                    index = cursor;
                    return true;
                }
            }

            return result;
        }
    }
}
