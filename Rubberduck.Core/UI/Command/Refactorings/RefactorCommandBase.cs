using System;
using System.Windows.Forms;
using NLog;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Resources;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    public abstract class RefactorCommandBase : CommandBase
    {
        protected readonly IRewritingManager RewritingManager;
        protected readonly ISelectionService SelectionService;

        protected RefactorCommandBase(IRewritingManager rewritingManager, ISelectionService selectionService)
            : base (LogManager.GetCurrentClassLogger())
        {
            RewritingManager = rewritingManager;
            SelectionService = selectionService;
        }

        protected void HandleInvalidSelection(object sender, EventArgs e)
        {
            System.Windows.Forms.MessageBox.Show(RubberduckUI.ExtractMethod_InvalidSelectionMessage, RubberduckUI.ExtractMethod_Caption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }
    }
}

