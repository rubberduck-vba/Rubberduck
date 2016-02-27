using Rubberduck.Parsing.Grammar;
using System.Text.RegularExpressions;

namespace Rubberduck.Parsing.Symbols
{
    public static class CommentExtensions
    {
        public static string GetComment(this VBAParser.RemCommentContext remComment)
        {
            string rawComment = remComment.GetText();
            string bodyOnly = Regex.Replace(rawComment, ":?REM(.*)", "$1", RegexOptions.IgnoreCase).TrimStart();
            return bodyOnly;
        }

        public static string GetComment(this VBAParser.CommentContext comment)
        {
            string rawComment = comment.GetText();
            string bodyOnly = rawComment.Substring(1).TrimStart();
            return bodyOnly;
        }
    }
}
