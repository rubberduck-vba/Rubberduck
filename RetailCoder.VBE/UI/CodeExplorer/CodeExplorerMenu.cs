using System;
using System.Diagnostics;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.VBA;

namespace Rubberduck.UI.CodeExplorer
{
    public class CodeExplorerMenu
    {
        private readonly VBE _vbe;
        private readonly AddIn _addin;
        private readonly IRubberduckParser _parser;

        private readonly CodeExplorerWindow _window;

        public CodeExplorerMenu(VBE vbe, AddIn addin, IRubberduckParser parser)
        {
            _vbe = vbe;
            _addin = addin;
            _parser = parser;

            _window = new CodeExplorerWindow();
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
            var presenter = new CodeExplorerDockablePresenter(_parser, _vbe, _addin, _window);
            presenter.Show();
        }
    }
}
