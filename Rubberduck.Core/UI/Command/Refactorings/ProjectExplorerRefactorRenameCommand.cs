using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI.Refactorings.Rename;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class ProjectExplorerRefactorRenameCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IRewritingManager _rewritingManager;
        private readonly IMessageBox _msgBox;

        public ProjectExplorerRefactorRenameCommand(IVBE vbe, RubberduckParserState state, IMessageBox msgBox, IRewritingManager rewritingManager) 
            : base (vbe)
        {
            _state = state;
            _rewritingManager = rewritingManager;
            _msgBox = msgBox;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            return _state.Status == ParserState.Ready;
        }

        protected override void OnExecute(object parameter)
        {
            using (var view = new RenameDialog(new RenameViewModel(_state)))
            {
                var factory = new RenamePresenterFactory(Vbe, view, _state);
                var refactoring = new RenameRefactoring(Vbe, factory, _msgBox, _state, _state.ProjectsProvider, _rewritingManager);

                var target = GetTarget();

                if (target != null)
                {
                    refactoring.Refactor(target);
                }
            }
        }

        private Declaration GetTarget()
        {
            if (Vbe.SelectedVBComponent == null)
            {
                return
                    _state.AllUserDeclarations.SingleOrDefault(d =>
                            d.DeclarationType == DeclarationType.Project && d.IdentifierName == Vbe.ActiveVBProject.Name);
            }
            
            return _state.AllUserDeclarations.SingleOrDefault(
                    t => t.IdentifierName == Vbe.SelectedVBComponent.Name &&
                            t.ProjectId == Vbe.ActiveVBProject.ProjectId &&
                            new[]
                                {
                                    DeclarationType.ClassModule,
                                    DeclarationType.Document,
                                    DeclarationType.ProceduralModule,
                                    DeclarationType.UserForm
                                }.Contains(t.DeclarationType));
        }
    }
}
