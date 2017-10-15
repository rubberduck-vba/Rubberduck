using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System.Collections.Generic;

namespace Rubberduck.Parsing.VBA
{
    public interface IProjectManager
    {
        IReadOnlyCollection<IVBProject> Projects { get; }

        void RefreshProjects();
        IReadOnlyCollection<QualifiedModuleName> AllModules();
    }
}
