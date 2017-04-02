using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using System.Linq;

namespace Rubberduck.Refactorings.Rename
{
    public class RenamePropertyHandler : RenameHandlerBase
    {
        public RenamePropertyHandler(RenameModel model, IMessageBox messageBox)
            : base(model, messageBox) { }

        override public string ErrorMessage
        {
            get
            {
                return string.Format(RubberduckUI.RenameDialog_PropertyRenameError, Model.Target.IdentifierName);
            }
        }

        override public void Rename()
        {
            // properties can have more than 1 member.
            var members = Model.Declarations.Named(Model.Target.IdentifierName)
                .Where(item => item.ProjectId == Model.Target.ProjectId
                    && item.ComponentName == Model.Target.ComponentName
                    && item.DeclarationType.HasFlag(DeclarationType.Property)).ToList();

            members.ForEach(member => RenameUsages(member));

            RenameDeclaration(Model.Target);

            Rewrite();

            Model.State.OnParseRequested(this);
        }
    }
}
