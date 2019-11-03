using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI.Command.Refactorings.Notifiers;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class FormDesignerRefactorRenameCommand : RefactorDeclarationCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IVBE _vbe;

        public FormDesignerRefactorRenameCommand(RenameRefactoring refactoring, RenameFailedNotifier renameFailedNotifier, IVBE vbe, RubberduckParserState state) 
            : base (refactoring, renameFailedNotifier, state)
        {
            _state = state;
            _vbe = vbe;

            AddToCanExecuteEvaluation(SpecializedEvaluateCanExecute);
        }

        private bool SpecializedEvaluateCanExecute(object parameter)
        {
            var target = GetTarget();
            return target != null && !_state.IsNewOrModified(target.QualifiedModuleName);
        }

        protected override Declaration GetTarget()
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
