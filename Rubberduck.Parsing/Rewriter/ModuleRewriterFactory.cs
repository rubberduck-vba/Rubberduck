using Antlr4.Runtime;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;

namespace Rubberduck.Parsing.Rewriter
{
    //We use a concrete implementation here instead of the CW auto-magic factories
    //because having to release the the rewriters later on is impractical
    //since they are stored in a different place than they get created
    //and do not require disposal themselves. 
    public class ModuleRewriterFactory : IModuleRewriterFactory
    {

        private readonly IProjectsProvider _projectsProvider;
        private readonly ISourceCodeHandler _sourceCodeHandler;

        public ModuleRewriterFactory(IProjectsProvider projectsProvider, ISourceCodeHandler sourceCodeHandler)
        {
            _projectsProvider = projectsProvider;
            _sourceCodeHandler = sourceCodeHandler;
        }

        public IModuleRewriter CodePaneRewriter(QualifiedModuleName module, ITokenStream tokenStream)
        {
            return  new CodePaneRewriter(module, tokenStream, _projectsProvider);
        }

        public IModuleRewriter AttributesRewriter(QualifiedModuleName module, ITokenStream tokenStream)
        {
            return new AttributesRewriter(module, tokenStream, _projectsProvider, _sourceCodeHandler);
        }
    }
}
