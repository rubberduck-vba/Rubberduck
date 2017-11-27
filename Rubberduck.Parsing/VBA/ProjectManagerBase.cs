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

        protected ProjectManagerBase(
            RubberduckParserState state,
            IVBE vbe)
        {
            _state = state ?? throw new ArgumentNullException(nameof(state));

            _vbe = vbe ?? throw new ArgumentNullException(nameof(vbe));
        }

        public abstract IReadOnlyCollection<QualifiedModuleName> AllModules();

        public IReadOnlyCollection<IVBProject> Projects => _state.Projects.AsReadOnly();

        public void RefreshProjects()
        {
            _state.RefreshProjects(_vbe);
        }
    }
}
