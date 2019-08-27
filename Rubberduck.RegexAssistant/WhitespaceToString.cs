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
                    spelledOutWhitespace = RubberduckUI.RegexAssistant_SpelledOut_Tab;
                    break;
                case " ":
                    spelledOutWhitespace = RubberduckUI.RegexAssistant_SpelledOut_Space;
                    break;
                case "\n":
                    spelledOutWhitespace = RubberduckUI.RegexAssistant_SpelledOut_NewLine;
                    break;
                case "\r":
                    spelledOutWhitespace = RubberduckUI.RegexAssistant_SpelledOut_CarriageReturn;
                    break;
                case "\r\n":
                    spelledOutWhitespace = RubberduckUI.RegexAssistant_SpelledOut_CarriageReturnNewLine;
                    break;
                default:
                    spelledOutWhitespace = RubberduckUI.RegexAssistant_SpelledOut_UnidentifiedWhitespace;
                    break;  
            }

            return string.Format(RubberduckUI.RegexAssistant_EncloseWhitespace_EnclosingFormat, spelledOutWhitespace);
        }
    }
}
