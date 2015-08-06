using System;
using System.Linq;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.UI.CodeExplorer;

namespace Rubberduck.UI.Command
{
    public class CodeExplorerCommand : ICommand
    {
        private readonly CodeExplorerDockablePresenter _presenter;

        public CodeExplorerCommand(CodeExplorerDockablePresenter presenter)
        {
            _presenter = presenter;
        }

        public void Execute()
        {
            _presenter.Show();
        }
    }

    public class CodeExplorerCommandMenuItem : CommandMenuItemBase
    {
        public CodeExplorerCommandMenuItem(ICommand command) 
            : base(command)
        {
        }

        public override string Key { get { return "RubberduckMenu_CodeExplorer"; } }
        public override int DisplayOrder { get { return (int)RubberduckMenuItemDisplayOrder.CodeExplorer; } }
    }
}