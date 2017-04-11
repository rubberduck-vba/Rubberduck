using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.VBA
{
    public class ProjectManager : IProjectManager
    {
        private readonly RubberduckParserState _state;
        private readonly IVBE _vbe;

        public ProjectManager(
            RubberduckParserState state,
            IVBE vbe)
        {
            if (state == null)
            {
                throw new ArgumentNullException(nameof(state));
            }
            if (vbe == null)
            {
                throw new ArgumentNullException(nameof(vbe));
            }

            _state = state;
            _vbe = vbe;
        }


        public IReadOnlyCollection<IVBProject> Projects
        {
            get
            {
                return _state.Projects.AsReadOnly();
            }
        }

        public IReadOnlyCollection<QualifiedModuleName> AllModules()
        {
            return Projects.SelectMany(project => project.VBComponents)
                            .Select(component => new QualifiedModuleName(component))
                            .ToHashSet()
                            .AsReadOnly(); ;
        }


        public void RefreshProjects()
        {
            _state.RefreshProjects(_vbe);
        }
    }
}
