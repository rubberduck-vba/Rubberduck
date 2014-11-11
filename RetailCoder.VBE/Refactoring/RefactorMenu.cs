using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.Reflection.VBA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.UnitTesting.UI
{
    internal class RefactorMenu : IDisposable
    {
        private readonly VBE _vbe;
        private readonly Parser _parser;

        public RefactorMenu(VBE vbe)
        {
            _vbe = vbe;
            _parser = new Parser();
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
            try
            {
                var tree = _parser.Parse(_vbe.ActiveCodePane.CodeModule);
            }
            catch(Exception exception)
            {

            }
        }

        private CommandBarButton AddMenuButton(CommandBarPopup menu)
        {
            return menu.Controls.Add(Type: MsoControlType.msoControlButton) as CommandBarButton;
        }

        public void Dispose()
        {
            
        }
    }
}
