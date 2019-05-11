using Antlr4.Runtime;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SourceCodeHandling;

namespace Rubberduck.Parsing.Rewriter
{
    //We use a concrete implementation here instead of the CW auto-magic factories
    //because having to release the the rewriters later on is impractical
    //since they are stored in a different place than they get created
    //and do not require disposal themselves. 
    public class ModuleRewriterFactory : IModuleRewriterFactory
    {

        private readonly ISourceCodeHandler _codePaneSourceCodeHandler;
        private readonly ISourceCodeHandler _attributesSourceCodeHandler;

        public ModuleRewriterFactory(ISourceCodeHandler codePaneSourceCodeHandler, ISourceCodeHandler attributesSourceCodeHandler)
        {
            _codePaneSourceCodeHandler = codePaneSourceCodeHandler;
            _attributesSourceCodeHandler = attributesSourceCodeHandler;
        }

        public IExecutableModuleRewriter CodePaneRewriter(QualifiedModuleName module, ITokenStream tokenStream)
        {
            return new ModuleRewriter(module, tokenStream, _codePaneSourceCodeHandler);
        }

        public IExecutableModuleRewriter AttributesRewriter(QualifiedModuleName module, ITokenStream tokenStream)
        {
            return new ModuleRewriter(module, tokenStream, _attributesSourceCodeHandler);
        }
    }
}
