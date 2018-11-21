using Antlr4.Runtime;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Rewriter
{
    public class StateTokenStreamCache : ITokenStreamCache
    {
        private readonly RubberduckParserState _state;

        public StateTokenStreamCache(RubberduckParserState state)
        {
            _state = state;
        }


        public ITokenStream CodePaneTokenStream(QualifiedModuleName module)
        {
            return _state.GetCodePaneTokenStream(module);
        }

        public ITokenStream AttributesTokenStream(QualifiedModuleName module)
        {
            return _state.GetAttributesTokenStream(module);
        }
    }
}