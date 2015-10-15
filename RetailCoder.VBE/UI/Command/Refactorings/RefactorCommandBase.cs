using Microsoft.Vbe.Interop;
using Rubberduck.VBEditor;
using System;
using System.Windows.Forms;
using Rubberduck.UI.ParserProgress;

namespace Rubberduck.UI.Command.Refactorings
{
    public abstract class RefactorCommandBase : CommandBase
    {
        protected readonly ParsingProgressPresenter ParserProgress;
        protected readonly IActiveCodePaneEditor Editor;
        protected readonly VBE Vbe;

        protected RefactorCommandBase(VBE vbe, ParsingProgressPresenter parserProgress, IActiveCodePaneEditor editor)
        {
            Vbe = vbe;
            ParserProgress = parserProgress;
            Editor = editor;
        }

        protected void HandleInvalidSelection(object sender, EventArgs e)
        {
            System.Windows.Forms.MessageBox.Show(RubberduckUI.ExtractMethod_InvalidSelectionMessage, RubberduckUI.ExtractMethod_Caption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }
    }
}

