using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Parsing.VBA
{
    public interface IProjectManager
    {
        ICollection<IVBProject> Projects { get; }

        void RefreshProjects();
        ICollection<QualifiedModuleName> AllModules();
    }
}
