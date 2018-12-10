using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.VBA
{
    public abstract class StateProjectManagerBase : IProjectManager
    {
        private readonly RubberduckParserState _state;
        private readonly IVBE _vbe;

        protected StateProjectManagerBase(
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


        public abstract IReadOnlyCollection<QualifiedModuleName> AllModules();


        public IReadOnlyCollection<(string ProjectId, IVBProject Project)> Projects
        {
            get
            {
                return _state.Projects.Select(project => (project.ProjectId, project)).ToList().AsReadOnly();
            }
        }


        public void RefreshProjects()
        {
            _state.RefreshProjects();
        }
    }
}
