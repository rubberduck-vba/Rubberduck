namespace Rubberduck.Parsing.VBA.Parsing.ParsingExceptions
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
