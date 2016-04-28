using Microsoft.Vbe.Interop;
using System;
using System.Windows.Forms;

namespace Rubberduck.UI.Command.Refactorings
{
    public abstract class RefactorCommandBase : CommandBase
    {
        protected readonly VBE Vbe;

        protected RefactorCommandBase(VBE vbe)
        {
            Vbe = vbe;
        }

        protected void HandleInvalidSelection(object sender, EventArgs e)
        {
            System.Windows.Forms.MessageBox.Show(RubberduckUI.ExtractMethod_InvalidSelectionMessage, RubberduckUI.ExtractMethod_Caption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }
    }
}

