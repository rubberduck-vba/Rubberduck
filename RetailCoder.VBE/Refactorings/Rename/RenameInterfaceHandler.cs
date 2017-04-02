using Rubberduck.Common;
using Rubberduck.UI;
using System.Linq;

namespace Rubberduck.Refactorings.Rename
{
    class RenameInterfaceHandler : RenameHandlerBase
    {
        public RenameInterfaceHandler(RenameModel model, IMessageBox messageBox)
            : base(model, messageBox) { }

        override public string ErrorMessage
        {
            get
            {
                return string.Format(RubberduckUI.RenameDialog_InterfaceRenameError, Model.Target.IdentifierName);
            }
        }

        override public void Rename()
        {
            RenameUsages(Model.Target);

            var implementations = Model.Declarations.FindInterfaceImplementationMembers()
                .Where(m => m.IdentifierName == Model.Target.ComponentName + '_' + Model.Target.IdentifierName)
                    .OrderByDescending(m => m.Selection.StartColumn).ToList();

            var newMemberName = Model.Target.ComponentName + '_' + Model.NewName;
            implementations.ForEach(imp => RenameDeclaration(imp, newMemberName));

            RenameDeclaration(Model.Target);

            Rewrite();

            Model.State.OnParseRequested(this);
        }
    }
}
