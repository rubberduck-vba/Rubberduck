using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Rename;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class ProjectExplorerRefactorRenameCommand : RefactorDeclarationCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _msgBox;
        private readonly IVBE _vbe;

        public ProjectExplorerRefactorRenameCommand(RenameRefactoring refactoring, IVBE vbe, RubberduckParserState state, IMessageBox msgBox, ISelectionService selectionService)
            : base(refactoring, state)
        {
            _state = state;
            _msgBox = msgBox;
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
