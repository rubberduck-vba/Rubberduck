using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.VBEditor;
using Rubberduck.UI.Command;
using System;

namespace Rubberduck.UI.Command.Refactorings
{
    public class RefactorCommandBase : CommandBase
    {
        protected readonly IRubberduckParser _parser;
        protected readonly IActiveCodePaneEditor _editor;
        protected readonly VBE _ide;

        public RefactorCommandBase(VBE ide, IRubberduckParser parser, IActiveCodePaneEditor editor)
        {
            _ide = ide;
            _parser = parser;
            _editor = editor;
        }

        void refactoring_InvalidSelection(object sender, EventArgs e)
        {
            System.Windows.Forms.MessageBox.Show(RubberduckUI.ExtractMethod_InvalidSelectionMessage, RubberduckUI.ExtractMethod_Caption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }
    }
}

