using System;
using System.Linq;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.UI.CodeExplorer;

namespace Rubberduck.UI.Command
{
    public class CodeExplorerCommand : ICommand, IDisposable
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

        public void Dispose()
        {
            _presenter.Dispose();
        }
    }
}