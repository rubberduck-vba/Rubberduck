using System.IO;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.VBEditor.SourceCodeHandling
{
    public class SourceFileHandlerSourceCodeHandlerAdapter : ISourceCodeHandler
    {
        private readonly IProjectsProvider _projectsProvider;
        private readonly ITempSourceFileHandler _tempSourceFileHandler;

        public SourceFileHandlerSourceCodeHandlerAdapter(ITempSourceFileHandler tempSourceFileHandler, IProjectsProvider projectsProvider)
        {
            _projectsProvider = projectsProvider;
            _tempSourceFileHandler = tempSourceFileHandler;
        }

        public string SourceCode(QualifiedModuleName module)
        {
            var component = _projectsProvider.Component(module);
            if (component == null)
            {
                return string.Empty;
            }

            return _tempSourceFileHandler.Read(component) ?? string.Empty;
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

            var fileName = _tempSourceFileHandler.Export(component);
            if (fileName == null || !File.Exists(fileName))
            {
                return;
            }
            File.WriteAllText(fileName, newCode);
            _tempSourceFileHandler.ImportAndCleanUp(component, fileName);

        }
    }
}
