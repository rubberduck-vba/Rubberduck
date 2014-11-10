using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using RetailCoderVBE.Reflection.VBA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RetailCoderVBE.UnitTesting.UI
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
        
        public void Initialize()
        {
            var menuBarControls = _vbe.CommandBars[1].Controls;
            var beforeIndex = FindMenuInsertionIndex(menuBarControls);
            var menu = menuBarControls.Add(Type: MsoControlType.msoControlPopup, Before: beforeIndex, Temporary: true) as CommandBarPopup;
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

        private int FindMenuInsertionIndex(CommandBarControls controls)
        {
            for (int i = 1; i <= controls.Count; i++)
            {
                // insert menu before "Window" built-in menu:
                if (controls[i].BuiltIn && controls[i].Caption == "&Window")
                {
                    return i;
                }
            }

            return controls.Count;
        }

        public void Dispose()
        {
            
        }
    }
}
