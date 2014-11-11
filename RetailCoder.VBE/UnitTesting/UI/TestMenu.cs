using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;

namespace Rubberduck.UnitTesting.UI
{
    internal class TestMenu : IDisposable
    {
        // 2743: play icon with stopwatch
        // 3039: module icon || 3119 || 621 || 589 || 472
        // 3170: class module icon

        private readonly VBE _vbe;
        private readonly TestEngine _engine;

        public TestMenu(VBE vbe)
        {
            _vbe = vbe;
            _engine = new TestEngine(_vbe);
        }

        private CommandBarButton _runAllTestsButton;
        public CommandBarButton RunAllTestsButton { get { return _runAllTestsButton; } }

        private CommandBarButton _windowsTestExplorerButton;
        public CommandBarButton WindowsTestExplorerButton { get { return _windowsTestExplorerButton; } }

        public void Initialize(CommandBarControls menuControls)
        {
            var menu = menuControls.Add(Type: MsoControlType.msoControlPopup, Temporary: true) as CommandBarPopup;
            menu.Caption = "Te&st";

            _windowsTestExplorerButton = AddMenuButton(menu);
            _windowsTestExplorerButton.Caption = "&Test Explorer";
            _windowsTestExplorerButton.FaceId = 3170;
            _windowsTestExplorerButton.Click += OnTestExplorerButtonClick;

            _runAllTestsButton = AddMenuButton(menu);
            _runAllTestsButton.BeginGroup = true;
            _runAllTestsButton.Caption = "&Run All Tests";
            _runAllTestsButton.FaceId = 186; // a "play" icon
            _runAllTestsButton.Click += OnRunAllTestsButtonClick;
        }

        private CommandBarButton AddMenuButton(CommandBarPopup menu)
        {
            return menu.Controls.Add(Type: MsoControlType.msoControlButton) as CommandBarButton;
        }

        void OnRunAllTestsButtonClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            _engine.SynchronizeTests();
            _engine.Run();
        }

        void OnTestExplorerButtonClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            _engine.ShowExplorer();
        }

        public void Dispose()
        {
            _engine.Dispose();
        }
    }
}
