using System.IO;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.VBEditor.SourceCodeHandling
{
    public class SourceFileHandlerSourceCodeHandlerAdapter : ISourceCodeHandler
    {
        private readonly IProjectsProvider _projectsProvider;
        private readonly ISourceFileHandler _sourceFileHandler;

        public SourceFileHandlerSourceCodeHandlerAdapter(ISourceFileHandler sourceFileHandler, IProjectsProvider projectsProvider)
        {
            _projectsProvider = projectsProvider;
            _sourceFileHandler = sourceFileHandler;
        }

        public string SourceCode(QualifiedModuleName module)
        {
            var component = _projectsProvider.Component(module);
            if (component == null)
            {
                return string.Empty;
            }

            return _sourceFileHandler.Read(component);
        }

        public void SubstituteCode(QualifiedModuleName module, string newCode)
        {
            if (module.ComponentType == ComponentType.Document)
            {
                //We cannot substitute the code of a document module via the file.
                return;
            }

            var component = _projectsProvider.Component(module);
            if (component == null)
            {
                return;
            }

            var file = _sourceFileHandler.Export(component);
            File.WriteAllText(file, newCode);
            _sourceFileHandler.Import(component, file);
        }
    }
}
