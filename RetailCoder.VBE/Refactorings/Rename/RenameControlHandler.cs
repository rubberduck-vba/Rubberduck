using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;

namespace Rubberduck.Refactorings.Rename
{
    class RenameControlHandler : RenameHandlerBase
    {
        public RenameControlHandler(RenameModel model, IMessageBox messageBox)
            : base(model, messageBox) { }

        override public string ErrorMessage
        {
            get
            {
                return string.Format(RubberduckUI.RenameDialog_ControlRenameError, Model.Target.IdentifierName);
            }
        }

        override public void Rename()
        {
            var controlDeclaration = Model.Target;

            Debug.Assert(Model.Target.DeclarationType == DeclarationType.Control, "Resolving User Selection to Declaration Failed");

            var module = Model.Target.QualifiedName.QualifiedModuleName.Component.CodeModule;
            var component = module.Parent;
            var control = component.Controls.SingleOrDefault(item => item.Name == controlDeclaration.IdentifierName);

            if (control == null)
            {
                MessageBox.Show(ErrorMessage, RubberduckUI.RenameDialog_Caption, MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);
                return;
            }

            var newControlName = Model.NewName;

            Debug.Assert(!newControlName.Equals(controlDeclaration.IdentifierName), "input validation fail: attempted to rename without changing the name");

            RenameUsages(controlDeclaration, newControlName);

            var handlers = Model.Declarations.FindEventHandlers(controlDeclaration).OrderByDescending(h => h.Selection.StartColumn).ToList();

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

            Model.State.OnParseRequested(this);
        }
    }
}
