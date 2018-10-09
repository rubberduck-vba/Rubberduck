using System;
using Antlr4.Runtime;

namespace Rubberduck.Parsing.VBA.Parsing.ParsingExceptions
{
    /// <summary>
    /// An exception that is thrown when the parser encounters a syntax error during one of two parses of an entire module.
    /// This exception indicates either a bug in the grammar... or non-compilable VBA code.
    /// </summary>
    [Serializable]
    public class ParsePassSyntaxErrorException : SyntaxErrorException
    {
        public ParsePassSyntaxErrorException(ParsePassSyntaxErrorInfo info)
            : this(info.Message, info.Exception, info.OffendingSymbol, info.LineNumber, info.Position, info.ModuleName, info.CodeKind) { }

        public ParsePassSyntaxErrorException(string message, RecognitionException innerException, IToken offendingSymbol, int line, int position, string moduleName, CodeKind codeKind)
            : base(message, innerException, offendingSymbol, line, position, codeKind)
        {
            ModuleName = moduleName;
            CodeKind = codeKind;
        }

        public string ModuleName { get; }
        public CodeKind CodeKind { get; }

        public override string ToString()
        {
            var parsePassText = CodeKind == CodeKind.CodePaneCode ? "code pane" : "exported";
            var exceptionText = 
$@"{base.ToString()}
Component: {ModuleName} ({parsePassText} version)";
            return exceptionText;
        }
    }
}
