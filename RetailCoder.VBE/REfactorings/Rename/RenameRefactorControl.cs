using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.Refactorings.Rename
{
    class RenameRefactorControl : RenameRefactorBase
    {
        public RenameRefactorControl(RubberduckParserState state)
            : base(state) { _errorMessage = string.Empty; }

        private string _errorMessage;
        public override string ErrorMessage => _errorMessage;

        public override void Rename(Declaration controlDeclaration, string newControlName)
        {
            _errorMessage = string.Format(RubberduckUI.RenameDialog_ControlRenameError, controlDeclaration.IdentifierName);

            Debug.Assert(controlDeclaration.DeclarationType == DeclarationType.Control, "Resolving User Selection to Declaration Failed");

            Debug.Assert(!newControlName.Equals(controlDeclaration.IdentifierName), "input validation fail: attempted to rename without changing the name");

            var module = controlDeclaration.QualifiedName.QualifiedModuleName.Component.CodeModule;
            var component = module.Parent;
            var control = component.Controls.SingleOrDefault(item => item.Name == controlDeclaration.IdentifierName);

            Debug.Assert(control != null, "input validation fail: unable to locate control in Controls collection");

            RenameUsages(controlDeclaration, newControlName);

            var handlers = State.AllUserDeclarations.FindEventHandlers(controlDeclaration).OrderByDescending(h => h.Selection.StartColumn).ToList();

            foreach (var handler in handlers)
            {
                var newMemberName = handler.IdentifierName.Replace(control.Name + '_', newControlName + '_');
                RenameUsages(handler, newMemberName);
            }

            foreach (var handler in handlers)
            {
                var newMemberName = handler.IdentifierName.Replace(control.Name + '_', newControlName + '_');
                RenameDeclaration(handler, newMemberName);
            }

            Rewrite();

            control.Name = newControlName;
        }
    }
}
