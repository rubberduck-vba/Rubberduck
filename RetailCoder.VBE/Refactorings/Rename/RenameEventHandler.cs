using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.Refactorings.Rename
{
    public class RenameEventHandler : RenameHandlerBase
    {
        public RenameEventHandler(RenameModel model, IMessageBox messageBox)
            : base(model, messageBox) { }

        override public string ErrorMessage
        {
            get
            {
                return string.Format(RubberduckUI.RenameDialog_EventRenameError, Model.Target.IdentifierName);
            }
        }

        override public void Rename()
        {
            Debug.Assert(Model.Target.DeclarationType == DeclarationType.Event, "Resolving User Selection to Declaration Failed");

            var eventDeclaration = Model.Target;

            RenameUsages(eventDeclaration, Model.NewName);

            var handlers = Model.Declarations.FindHandlersForEvent(eventDeclaration).ToList();
            handlers.ForEach(handler => RenameDeclaration(handler.Item2, handler.Item1.IdentifierName + '_' + Model.NewName));

            RenameDeclaration(eventDeclaration, Model.NewName);

            Rewrite();

            Model.State.OnParseRequested(this);
        }
    }
}
