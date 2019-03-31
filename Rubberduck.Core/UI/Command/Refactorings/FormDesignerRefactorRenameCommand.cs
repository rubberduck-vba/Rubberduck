using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Rename;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class FormDesignerRefactorRenameCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;
        private readonly IRefactoringPresenterFactory _factory;
        private readonly IVBE _vbe;

        public FormDesignerRefactorRenameCommand(IVBE vbe, RubberduckParserState state, IMessageBox messageBox, IRefactoringPresenterFactory factory, IRewritingManager rewritingManager, ISelectionService selectionService) 
            : base (rewritingManager, selectionService)
        {
            _state = state;
            _messageBox = messageBox;
            _factory = factory;
            _vbe = vbe;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            if (_state.Status != ParserState.Ready)
            {
                return false;
            }

            var target = GetTarget();
            return target != null && !_state.IsNewOrModified(target.QualifiedModuleName);
        }

        protected override void OnExecute(object parameter)
        {
            var refactoring = new RenameRefactoring(_factory, _messageBox, _state, _state.ProjectsProvider, RewritingManager, SelectionService);
            var target = GetTarget();
            if (target != null)
            {
                refactoring.Refactor(target);
            }
            
        }

        private Declaration GetTarget()
        {
            string projectId;
            using (var activeProject = _vbe.ActiveVBProject)
            {
                projectId = activeProject.ProjectId;
            }

            using (var component = _vbe.SelectedVBComponent)
            {
                if (!(component?.HasDesigner ?? false))
                {
                    return null;
                }

                DeclarationType selectedType;
                string selectedName;
                using (var selectedControls = component.SelectedControls)
                {
                    var selectedCount = selectedControls.Count;
                    if (selectedCount > 1)
                    {
                        return null;
                    }

                    // Cannot use DeclarationType.UserForm, parser only assigns UserForms the ClassModule flag
                    (selectedType, selectedName) = selectedCount == 0
                        ? (DeclarationType.ClassModule, component.Name)
                        : (DeclarationType.Control, selectedControls[0].Name);
                }

                return _state.DeclarationFinder
                    .MatchName(selectedName)
                    .SingleOrDefault(m => m.ProjectId == projectId
                                          && m.DeclarationType.HasFlag(selectedType)
                                          && m.ComponentName == component.Name);
            }
        }

        private Declaration GetTarget(QualifiedModuleName qualifiedModuleName)
        {
            var projectId = qualifiedModuleName.ProjectId;
            var component = _state.ProjectsProvider.Component(qualifiedModuleName);

            if (component?.HasDesigner ?? false)
            {
                return _state.DeclarationFinder
                    .MatchName(qualifiedModuleName.Name)
                    .SingleOrDefault(m => m.ProjectId == projectId
                                          && m.DeclarationType.HasFlag(qualifiedModuleName.ComponentType)
                                          && m.ComponentName == component.Name);
            }
            return null;
        }
    }
}
