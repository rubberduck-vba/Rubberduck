using System;
using Antlr4.Runtime;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Parsing.Symbols.ParsingExceptions
{
    /// <summary>
    /// An exception that is thrown when the parser encounters a syntax error while preprocessing a module.
    /// This exception indicates either a bug in the grammar... or non-compilable VBA code.
    /// </summary>
    [Serializable]
    public class PreprocessorSyntaxErrorException : ParsePassSyntaxErrorException
    {
        public PreprocessorSyntaxErrorException(ParsePassSyntaxErrorInfo info)
            : this(info.Message, info.Exception, info.OffendingSymbol, info.LineNumber, info.Position, info.ComponentName, info.ParsePass) { }

        public PreprocessorSyntaxErrorException(string message, RecognitionException innerException, IToken offendingSymbol, int line, int position, string componentName, ParsePass parsePass)
            : base(message, innerException, offendingSymbol, line, position, componentName, parsePass)
        {}

        public override string ToString()
        {
            var exceptionText = 
$@"{base.ToString()}
ParseType: Preprocessing";
            return exceptionText;
        }
    }
}
