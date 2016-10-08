using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ExtractInterface;
using Rubberduck.UI.Refactorings;
using Rubberduck.VBEditor.SafeComWrappers.VBA.Abstract;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorExtractInterfaceCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;

        public RefactorExtractInterfaceCommand(IVBE vbe, RubberduckParserState state, IMessageBox messageBox)
            :base(vbe)
        {
            _state = state;
            _messageBox = messageBox;
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            var selection = Vbe.ActiveCodePane.GetQualifiedSelection();
            if (!selection.HasValue)
            {
                return false;
            }

            var target = _state.AllUserDeclarations.SingleOrDefault(item =>
                item.QualifiedName.QualifiedModuleName.Equals(selection.Value.QualifiedName)
                && item.IdentifierName == selection.Value.QualifiedName.ComponentName
                && (item.DeclarationType == DeclarationType.ClassModule || item.DeclarationType == DeclarationType.Document || item.DeclarationType == DeclarationType.UserForm));
            var hasMembers = _state.AllUserDeclarations.Any(item => item.DeclarationType.HasFlag(DeclarationType.Member) && item.ParentDeclaration != null && item.ParentDeclaration.Equals(target));

            // true if active code pane is for a class/document/form module
            return target != null && hasMembers;
        }

        protected override void ExecuteImpl(object parameter)
        {
            if (Vbe.ActiveCodePane == null)
            {
                return;
            }

            using (var view = new ExtractInterfaceDialog())
            {
                var factory = new ExtractInterfacePresenterFactory(Vbe, _state, view);
                var refactoring = new ExtractInterfaceRefactoring(Vbe, _state, _messageBox, factory);
                refactoring.Refactor();
            }
        }
    }
}
