using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ExtractInterface;
using Rubberduck.UI.Refactorings;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

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

        private static readonly IReadOnlyList<DeclarationType> ModuleTypes = new[] 
        {
            DeclarationType.ClassModule,
            DeclarationType.UserForm, 
            DeclarationType.ProceduralModule, 
        };

        protected override bool CanExecuteImpl(object parameter)
        {
            var selection = Vbe.ActiveCodePane.GetQualifiedSelection();
            if (!selection.HasValue)
            {
                return false;
            }

            var interfaceClass = _state.AllUserDeclarations.SingleOrDefault(item =>
                item.QualifiedName.QualifiedModuleName.Equals(selection.Value.QualifiedName)
                && ModuleTypes.Contains(item.DeclarationType));

            // interface class must have members to be implementable
            var hasMembers = _state.AllUserDeclarations.Any(item => 
                item.DeclarationType.HasFlag(DeclarationType.Member) 
                && item.ParentDeclaration != null 
                && item.ParentDeclaration.Equals(interfaceClass));

            // true if active code pane is for a class/document/form module
            return interfaceClass != null && hasMembers;
        }

        protected override void ExecuteImpl(object parameter)
        {
            if (Vbe.ActiveCodePane == null)
            {
                return;
            }

            using (var view = new ExtractInterfaceDialog(new ExtractInterfaceViewModel()))
            {
                var factory = new ExtractInterfacePresenterFactory(Vbe, _state, view);
                var refactoring = new ExtractInterfaceRefactoring(Vbe, _messageBox, factory);
                refactoring.Refactor();
            }
        }
    }
}
