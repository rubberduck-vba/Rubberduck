using System;
using System.Collections.Generic;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.VBA
{
    public abstract class ProjectManagerBase : IProjectManager
    {
        private readonly RubberduckParserState _state;
        private readonly IVBE _vbe;

        public ProjectManagerBase(
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


        public IReadOnlyCollection<IVBProject> Projects
        {
            get
            {
                return _state.Projects.AsReadOnly();
            }
        }


        public void RefreshProjects()
        {
            _state.RefreshProjects(_vbe);
        }
    }
}
