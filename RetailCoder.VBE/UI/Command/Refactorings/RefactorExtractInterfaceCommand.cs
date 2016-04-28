using System.Diagnostics;
using System.Linq;
using Microsoft.Vbe.Interop;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ExtractInterface;
using Rubberduck.UI.Refactorings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorExtractInterfaceCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;

        public RefactorExtractInterfaceCommand(VBE vbe, RubberduckParserState state, IMessageBox messageBox)
            :base(vbe)
        {
            _state = state;
            _messageBox = messageBox;
        }

        private static readonly vbext_ComponentType[] ModuleTypes =
        {
            vbext_ComponentType.vbext_ct_ClassModule, 
            vbext_ComponentType.vbext_ct_Document, 
            vbext_ComponentType.vbext_ct_MSForm, 
        };

        public override bool CanExecute(object parameter)
        {
            var activePane = Vbe.ActiveCodePane;
            if (activePane == null)
            {
                return false;
            }

            var selection = activePane.GetQualifiedSelection();
            var target = _state.AllUserDeclarations.SingleOrDefault(item =>
                item.QualifiedName.QualifiedModuleName.Equals(selection.QualifiedName)
                && item.IdentifierName == selection.QualifiedName.ComponentName
                && (item.DeclarationType == DeclarationType.ClassModule || item.DeclarationType == DeclarationType.Document || item.DeclarationType == DeclarationType.UserForm));
            var hasMembers = _state.AllUserDeclarations.Any(item => item.DeclarationType.HasFlag(DeclarationType.Member) && item.ParentDeclaration != null && item.ParentDeclaration.Equals(target));

            // true if active code pane is for a class/document/form module
            var canExecute = ModuleTypes.Contains(Vbe.ActiveCodePane.CodeModule.Parent.Type) && target != null && hasMembers;

            Debug.WriteLine("{0}.CanExecute evaluates to {1}", GetType().Name, canExecute);
            return canExecute;
        }

        public override void Execute(object parameter)
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