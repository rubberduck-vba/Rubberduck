using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI.Refactorings.Rename;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class FormDesignerRefactorRenameCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;

        public FormDesignerRefactorRenameCommand(IVBE vbe, RubberduckParserState state, IMessageBox messageBox) 
            : base (vbe)
        {
            _state = state;
            _messageBox = messageBox;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            var target = GetTarget();
            return _state.Status == ParserState.Ready && target != null && !_state.IsNewOrModified(target.QualifiedModuleName);
        }

        protected override void OnExecute(object parameter)
        {
            using (var view = new RenameDialog(new RenameViewModel(_state)))
            {
                var factory = new RenamePresenterFactory(Vbe, view, _state);
                var refactoring = new RenameRefactoring(Vbe, factory, _messageBox, _state);

                var target = GetTarget();

                if (target != null)
                {
                    refactoring.Refactor(target);
                }
            }
        }

        private Declaration GetTarget(QualifiedModuleName? qualifiedModuleName = null)
        {
            if (qualifiedModuleName.HasValue)
            {
                return GetTarget(qualifiedModuleName.Value);
            }

            string projectId;
            using (var activeProject = Vbe.ActiveVBProject)
            {
                projectId = activeProject.ProjectId;
            }

            using (var component = Vbe.SelectedVBComponent)
            {
                if (component?.HasDesigner ?? false)
                {
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

            return null;
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
