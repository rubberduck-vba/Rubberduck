using System.Linq;
using System.Runtime.InteropServices;
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
            return _state.Status == ParserState.Ready && GetTarget() != null;
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
            (var projectId, var component) = qualifiedModuleName.HasValue 
                ? (qualifiedModuleName.Value.ProjectId, qualifiedModuleName.Value.Component)
                : (Vbe.ActiveVBProject.ProjectId, Vbe.SelectedVBComponent);
                        
            if (component?.HasDesigner ?? false)
            {
                if (qualifiedModuleName.HasValue)
                {
                    return _state.DeclarationFinder
                        .MatchName(qualifiedModuleName.Value.Name)
                        .SingleOrDefault(m => m.ProjectId == projectId
                            && m.DeclarationType.HasFlag(qualifiedModuleName.Value.ComponentType)
                            && m.ComponentName == component.Name);
                }
                
                var selectedCount = component.SelectedControls.Count;
                if (selectedCount > 1) { return null; }
                
                // Cannot use DeclarationType.UserForm, parser only assigns UserForms the ClassModule flag
                (var selectedType, var selectedName) = selectedCount == 0
                    ? (DeclarationType.ClassModule, component.Name)
                    : (DeclarationType.Control, component.SelectedControls[0].Name);
                
                return _state.DeclarationFinder
                    .MatchName(selectedName)
                    .SingleOrDefault(m => m.ProjectId == projectId
                        && m.DeclarationType.HasFlag(selectedType)
                        && m.ComponentName == component.Name);
             }
            return null;
        }
    }
}
