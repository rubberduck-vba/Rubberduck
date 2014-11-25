using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.UI.CodeExplorer;
using Rubberduck.VBA.Parser;

namespace Rubberduck.UI.CodeExplorer
{
    class CodeExplorerMenu
    {
        private readonly VBE _vbe;
        private readonly AddIn _addin;
        private readonly Parser _parser;

        public CodeExplorerMenu(VBE vbe, AddIn addin)
        {
            _vbe = vbe;
            _addin = addin;
            _parser = new Parser();
        }

        private CommandBarButton _codeExplorerButton;
        public CommandBarButton CodeExplorerButton { get { return _codeExplorerButton; } }

        public void Initialize(CommandBarControls menuControls)
        {
            _codeExplorerButton = menuControls.Add(Type: MsoControlType.msoControlButton, Temporary: true) as CommandBarButton;
            _codeExplorerButton.Caption = "&Code Explorer";
            _codeExplorerButton.BeginGroup = true;

            _codeExplorerButton.FaceId = 3039;
            _codeExplorerButton.Click += OnCodeExplorerButtonClick;
        }

        private void OnCodeExplorerButtonClick(CommandBarButton button, ref bool CancelDefault)
        {
            var presenter = new CodeExplorerDockablePresenter(_parser, _vbe, _addin);
            presenter.Show();
        }
    }
}
