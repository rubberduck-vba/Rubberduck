using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SourceCodeHandling
{
    public class CodePaneSourceCodeHandler : ISourceCodeHandler
    {
        private readonly IProjectsProvider _projectsProvider;

        public CodePaneSourceCodeHandler(IProjectsProvider projectsProvider)
        {
            _projectsProvider = projectsProvider;
        }

        public string SourceCode(QualifiedModuleName module)
        {
            var component = _projectsProvider.Component(module);
            if (component == null)
            {
                return string.Empty;
            }

            using (var codeModule = component.CodeModule)
            {
                return codeModule.Content() ?? string.Empty;
            }
        }

        public void SubstituteCode(QualifiedModuleName module, string newCode)
        {
            var component = _projectsProvider.Component(module);
            if (component == null)
            {
                return;
            }

            using (var codeModule = component.CodeModule)
            {
                codeModule.Clear();
                codeModule.InsertLines(1, newCode);
            }
        }
    }
}
