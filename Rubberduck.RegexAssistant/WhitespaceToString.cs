using Rubberduck.Resources;
using System.Text.RegularExpressions;

namespace Rubberduck.RegexAssistant
{
    static class WhitespaceToString
    {
        private static readonly Regex whitespace = new Regex("\\s");
        public static bool IsFullySpellingOutApplicable(string value, out string spelledOutWhiteSpace)
        {
            if (!whitespace.IsMatch(value))
            {
                spelledOutWhiteSpace = string.Empty;
                return false;
            }

            spelledOutWhiteSpace = ConvertWhitespaceToStringName(value);
            return true;
        }

        private static string ConvertWhitespaceToStringName(string whiteSpace)
        {
            string spelledOutWhitespace;
            switch (whiteSpace)
            {
                case "\t":
                    spelledOutWhitespace = RegexAssistantUI.SpelledOut_Tab;
                    break;
                case " ":
                    spelledOutWhitespace = RegexAssistantUI.SpelledOut_Space;
                    break;
                case "\n":
                    spelledOutWhitespace = RegexAssistantUI.SpelledOut_NewLine;
                    break;
                case "\r":
                    spelledOutWhitespace = RegexAssistantUI.SpelledOut_CarriageReturn;
                    break;
                case "\r\n":
                    spelledOutWhitespace = RegexAssistantUI.SpelledOut_CarriageReturnNewLine;
                    break;
                default:
                    spelledOutWhitespace = RegexAssistantUI.SpelledOut_UnidentifiedWhitespace;
                    break;  
            }

            return string.Format(RegexAssistantUI.EncloseWhitespace_EnclosingFormat, spelledOutWhitespace);
        }
    }
}
