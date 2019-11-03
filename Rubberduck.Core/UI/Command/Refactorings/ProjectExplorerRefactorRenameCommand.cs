using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI.Command.Refactorings.Notifiers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class ProjectExplorerRefactorRenameCommand : RefactorDeclarationCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IVBE _vbe;

        public ProjectExplorerRefactorRenameCommand(RenameRefactoring refactoring, RenameFailedNotifier renameFailedNotifier, IVBE vbe, RubberduckParserState state)
            : base(refactoring, renameFailedNotifier, state)
        {
            _state = state;
            _vbe = vbe;
        }

        protected override Declaration GetTarget()
        {
            string selectedComponentName;
            using (var selectedComponent = _vbe.SelectedVBComponent)
            {
                selectedComponentName = selectedComponent?.Name;
            }

            string activeProjectId;
            using (var activeProject = _vbe.ActiveVBProject)
            {
                activeProjectId = activeProject?.ProjectId;
            }

            if (activeProjectId == null)
            {
                return null;
            }

            if (selectedComponentName == null)
            {
                return _state.DeclarationFinder.UserDeclarations(DeclarationType.Project)
                    .SingleOrDefault(d => d.ProjectId == activeProjectId);
            }

            return _state.DeclarationFinder.UserDeclarations(DeclarationType.Module)
                .SingleOrDefault(t => t.IdentifierName == selectedComponentName
                                      && t.ProjectId == activeProjectId);
        }
    }
}
