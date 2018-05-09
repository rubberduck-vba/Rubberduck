using System;
using System.Windows.Forms;
using NLog;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.Refactorings
{
    public abstract class RefactorCommandBase : CommandBase
    {
        protected readonly IVBE Vbe;

        protected RefactorCommandBase(IVBE vbe)
            : base (LogManager.GetCurrentClassLogger())
        {
            Vbe = vbe;
        }

        protected void HandleInvalidSelection(object sender, EventArgs e)
        {
            System.Windows.Forms.MessageBox.Show(RubberduckUI.ExtractMethod_InvalidSelectionMessage, RubberduckUI.ExtractMethod_Caption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }
    }
}

