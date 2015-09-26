using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.VBEditor;
using System;
using System.Windows.Forms;

namespace Rubberduck.UI.Command.Refactorings
{
    public abstract class RefactorCommandBase : CommandBase
    {
        protected readonly IRubberduckParser Parser;
        protected readonly IActiveCodePaneEditor Editor;
        protected readonly VBE Vbe;

        protected RefactorCommandBase(VBE vbe, IRubberduckParser parser, IActiveCodePaneEditor editor)
        {
            Vbe = vbe;
            Parser = parser;
            Editor = editor;
        }

        protected void refactoring_InvalidSelection(object sender, EventArgs e)
        {
            System.Windows.Forms.MessageBox.Show(RubberduckUI.ExtractMethod_InvalidSelectionMessage, RubberduckUI.ExtractMethod_Caption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }
    }
}

