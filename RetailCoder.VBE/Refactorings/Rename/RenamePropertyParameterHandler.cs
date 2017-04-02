using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using System.Linq;

namespace Rubberduck.Refactorings.Rename
{
    class RenamePropertyParameterHandler : RenameHandlerBase
    {
        public RenamePropertyParameterHandler(RenameModel model, IMessageBox messageBox)
            : base(model, messageBox) { }

        override public string ErrorMessage
        {
            get
            {
                return string.Format(RubberduckUI.RenameDialog_PropertyParameterRenameError, Model.Target.IdentifierName, Model.Target.ParentDeclaration.IdentifierName);
            }
        }

        override public void Rename()
        {
            var parameters = Model.Declarations.Where(d =>
                d.DeclarationType == DeclarationType.Parameter
                && d.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Property)
                && d.IdentifierName == Model.Target.IdentifierName).ToList();

            parameters.ForEach(param => RenameUsages(param));

            parameters.ForEach(param => RenameDeclaration(param));

            Rewrite();

            Model.State.OnParseRequested(this);
        }
    }
}
