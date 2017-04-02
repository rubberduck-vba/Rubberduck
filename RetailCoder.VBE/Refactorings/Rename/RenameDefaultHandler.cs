using Rubberduck.UI;

namespace Rubberduck.Refactorings.Rename
{
    public class RenameDefaultHandler :RenameHandlerBase
    {
        public RenameDefaultHandler(RenameModel model, IMessageBox messageBox)
            : base(model, messageBox) { }

        override public string ErrorMessage
        {
            get
            {
                return string.Format(RubberduckUI.RenameDialog_DefaultRenameError, Model.Target.IdentifierName);
            }
        }

        override public void Rename()
        {
            RenameUsages(Model.Target);

            RenameDeclaration(Model.Target);

            Rewrite();

            Model.State.OnParseRequested(this);
        }
    }
}
