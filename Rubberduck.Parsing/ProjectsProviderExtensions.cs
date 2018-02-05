using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing
{
    public static class ProjectsProviderExtensions
    {
        public static IVBComponent Component(this IProjectsProvider provider, Declaration declaration)
        {
            return provider.Component(declaration.QualifiedModuleName);
        }
    }
}
