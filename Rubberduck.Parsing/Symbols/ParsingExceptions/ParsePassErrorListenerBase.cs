using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols.ParsingExceptions
{
    public class ParsePassErrorListenerBase : RubberduckParseErrorListenerBase
    {
        protected string ModuleName { get; }

        public ParsePassErrorListenerBase(string moduleName, CodeKind codeKind) 
        :base(codeKind)
        {
            ModuleName = moduleName;
        }
    }
}
