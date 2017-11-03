using System;
using Antlr4.Runtime;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Parsing.Symbols.ParsingExceptions
{
    /// <summary>
    /// An exception that is thrown when the parser encounters a syntax error during one of two parses of an entire module.
    /// This exception indicates either a bug in the grammar... or non-compilable VBA code.
    /// </summary>
    [Serializable]
    public class ParsePassSyntaxErrorException : SyntaxErrorException
    {
        public ParsePassSyntaxErrorException(ParsePassSyntaxErrorInfo info)
            : this(info.Message, info.Exception, info.OffendingSymbol, info.LineNumber, info.Position, info.ComponentName, info.ParsePass) { }

        public ParsePassSyntaxErrorException(string message, RecognitionException innerException, IToken offendingSymbol, int line, int position, string componentName, ParsePass parsePass)
            : base(message, innerException, offendingSymbol, line, position)
        {
            ComponentName = componentName;
            ParsePass = parsePass;
        }

        public string ComponentName { get; }
        public ParsePass ParsePass { get; }

        public override string ToString()
        {
            var parsePassText = ParsePass == ParsePass.CodePanePass ? "code pane" : "exported";
            var exceptionText = 
$@"{base.ToString()}
Component: {ComponentName} ({parsePassText} version)";
            return exceptionText;
        }
    }
}
