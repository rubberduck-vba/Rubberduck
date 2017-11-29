using Rubberduck.Parsing.Grammar;
using System.Text.RegularExpressions;

namespace Rubberduck.Parsing.Symbols
{
    public static class CommentExtensions
    {
        public static string GetComment(this VBAParser.RemCommentContext remComment)
        {
            var rawComment = remComment.GetText();
            var bodyOnly = Regex.Replace(rawComment, ":?REM(.*)", "$1", RegexOptions.IgnoreCase).TrimStart();
            return bodyOnly;
        }

        public static string GetComment(this VBAParser.CommentContext comment)
        {
            var rawComment = comment.GetText();
            var bodyOnly = rawComment.Substring(1).TrimStart();
            return bodyOnly;
        }
    }
}
