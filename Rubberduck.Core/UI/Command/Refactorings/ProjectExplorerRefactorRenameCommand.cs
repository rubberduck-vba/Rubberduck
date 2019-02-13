using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Rename;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class ProjectExplorerRefactorRenameCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _msgBox;
        private readonly IRefactoringPresenterFactory _factory;
        private readonly IVBE _vbe;

        public ProjectExplorerRefactorRenameCommand(IVBE vbe, RubberduckParserState state, IMessageBox msgBox, IRefactoringPresenterFactory factory, IRewritingManager rewritingManager, ISelectionService selectionService)
            : base(rewritingManager, selectionService)
        {
            _state = state;
            _msgBox = msgBox;
            _factory = factory;
            _vbe = vbe;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            return _state.Status == ParserState.Ready;
        }

        protected override void OnExecute(object parameter)
        {
            var refactoring = new RenameRefactoring(_factory, _msgBox, _state, _state.ProjectsProvider, RewritingManager, SelectionService);
            var target = GetTarget();
            if (target != null)
            {
                refactoring.Refactor(target);
            }
        }

        private Declaration GetTarget()
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
