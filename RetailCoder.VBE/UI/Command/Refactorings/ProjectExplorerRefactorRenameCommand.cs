using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI.Refactorings;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class ProjectExplorerRefactorRenameCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly ICodePaneWrapperFactory _wrapperWrapperFactory;

        public ProjectExplorerRefactorRenameCommand(VBE vbe, RubberduckParserState state, ICodePaneWrapperFactory wrapperWrapperFactory) 
            : base (vbe)
        {
            _state = state;
            _wrapperWrapperFactory = wrapperWrapperFactory;
        }

        public override void Execute(object parameter)
        {
            using (var view = new RenameDialog())
            {
                var factory = new RenamePresenterFactory(Vbe, view, _state, new MessageBox(), _wrapperWrapperFactory);
                var refactoring = new RenameRefactoring(Vbe, factory, new MessageBox(), _state);

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
                            t.ProjectId == Vbe.ActiveVBProject.HelpFile &&
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