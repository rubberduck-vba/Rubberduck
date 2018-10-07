using System;
using Antlr4.Runtime;

namespace Rubberduck.Parsing.VBA.Parsing.ParsingExceptions
{
    /// <summary>
    /// An exception that is thrown when the parser encounters a syntax error while preprocessing a module.
    /// This exception indicates either a bug in the grammar... or non-compilable VBA code.
    /// </summary>
    [Serializable]
    public class PreprocessorSyntaxErrorException : ParsePassSyntaxErrorException
    {
        public PreprocessorSyntaxErrorException(ParsePassSyntaxErrorInfo info)
            : this(info.Message, info.Exception, info.OffendingSymbol, info.LineNumber, info.Position, info.ModuleName, info.CodeKind) { }

        public PreprocessorSyntaxErrorException(string message, RecognitionException innerException, IToken offendingSymbol, int line, int position, string moduleName, CodeKind codeKind)
            : base(message, innerException, offendingSymbol, line, position, moduleName, codeKind)
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
