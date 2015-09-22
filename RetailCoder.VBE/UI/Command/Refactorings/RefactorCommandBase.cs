using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.VBEditor;
using System;
using System.Windows.Forms;

namespace Rubberduck.UI.Command.Refactorings
{
    public abstract class RefactorCommandBase : CommandBase
    {
        protected readonly IRubberduckParser _parser;
        protected readonly IActiveCodePaneEditor _editor;
        protected readonly VBE _ide;

        protected RefactorCommandBase(VBE ide, IRubberduckParser parser, IActiveCodePaneEditor editor)
        {
            _ide = ide;
            _parser = parser;
            _editor = editor;
        }

        protected void refactoring_InvalidSelection(object sender, EventArgs e)
        {
            System.Windows.Forms.MessageBox.Show(RubberduckUI.ExtractMethod_InvalidSelectionMessage, RubberduckUI.ExtractMethod_Caption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }
    }
}

