using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.VBA.Parser;

namespace Rubberduck.UI.CodeExplorer
{
    [ComVisible(false)]
    public class CodeExplorerMenu
    {
        private readonly VBE _vbe;
        private readonly AddIn _addin;
        private readonly Parser _parser;

        public CodeExplorerMenu(VBE vbe, AddIn addin, Parser parser)
        {
            _vbe = vbe;
            _addin = addin;
            _parser = parser;
        }

        private CommandBarButton _codeExplorerButton;

        public void Initialize(CommandBarControls menuControls)
        {
            _codeExplorerButton = menuControls.Add(MsoControlType.msoControlButton, Temporary: true) as CommandBarButton;
            Debug.Assert(_codeExplorerButton != null);

            _codeExplorerButton.Caption = "&Code Explorer";
            _codeExplorerButton.BeginGroup = true;

            _codeExplorerButton.FaceId = 3039;
            _codeExplorerButton.Click += OnCodeExplorerButtonClick;
        }

        private void OnCodeExplorerButtonClick(CommandBarButton button, ref bool cancelDefault)
        {
            var presenter = new CodeExplorerDockablePresenter(_parser, _vbe, _addin);
            presenter.Show();
        }
    }
}
