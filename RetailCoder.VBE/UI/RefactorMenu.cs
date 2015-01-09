using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.UI.CodeExplorer;
using Rubberduck.VBA;

namespace Rubberduck.UI
{
    [ComVisible(false)]
    public class RefactorMenu : IDisposable
    {
        private readonly VBE _vbe;
        private readonly AddIn _addin;
        private readonly Parser _parser;

        public RefactorMenu(VBE vbe, AddIn addin, Parser parser)
        {
            _vbe = vbe;
            _addin = addin;
            _parser = parser;
        }

        private CommandBarButton _parseModuleButton;
        public CommandBarButton ParseModuleButton { get { return _parseModuleButton; } }
        
        public void Initialize(CommandBarControls menuControls)
        {
            var menu = menuControls.Add(Type: MsoControlType.msoControlPopup, Temporary: true) as CommandBarPopup;
            menu.Caption = "&Refactor";

            _parseModuleButton = AddMenuButton(menu);
            _parseModuleButton.Caption = "&Parse module";
            _parseModuleButton.FaceId = 3039;
            _parseModuleButton.Click += OnParseModuleButtonClick;

        }

        private void OnParseModuleButtonClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            using (var presenter = new CodeExplorerDockablePresenter(_parser, _vbe, _addin))
            {
                presenter.Show();
            }
        }

        private CommandBarButton AddMenuButton(CommandBarPopup menu)
        {
            return menu.Controls.Add(MsoControlType.msoControlButton) as CommandBarButton;
        }

        public void Dispose()
        {
            
        }
    }
}
