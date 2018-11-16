using Antlr4.Runtime;

namespace Rubberduck.Parsing.VBA.Parsing.ParsingExceptions
{
    public class ParsePassSyntaxErrorInfo : SyntaxErrorInfo
    {
        public ParsePassSyntaxErrorInfo(string message, RecognitionException innerException, IToken offendingSymbol, int line, int position, string moduleName, CodeKind codeKind)
        :base(message, innerException, offendingSymbol, line, position, codeKind){
            ModuleName = moduleName;
            CodeKind = codeKind;
        }

        public string ModuleName { get; }
        public CodeKind CodeKind { get; }
    }
}