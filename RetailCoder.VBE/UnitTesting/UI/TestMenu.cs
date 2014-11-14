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

        private readonly TestEngine _engine;

        public TestMenu(VBE vbe)
        {
            _engine = new TestEngine(vbe);
        }

        public CommandBarButton RunAllTestsButton { get; private set; }
        public CommandBarButton WindowsTestExplorerButton { get; private set; }

        public void Initialize(CommandBarControls menuControls)
        {
            var menu = menuControls.Add(Type: MsoControlType.msoControlPopup, Temporary: true) as CommandBarPopup;
            menu.Caption = "Te&st";

            WindowsTestExplorerButton = AddMenuButton(menu);
            WindowsTestExplorerButton.Caption = "&Test Explorer";
            WindowsTestExplorerButton.FaceId = 3170;
            WindowsTestExplorerButton.Click += OnTestExplorerButtonClick;

            RunAllTestsButton = AddMenuButton(menu);
            RunAllTestsButton.BeginGroup = true;
            RunAllTestsButton.Caption = "&Run All Tests";
            RunAllTestsButton.FaceId = 186; // a "play" icon
            RunAllTestsButton.Click += OnRunAllTestsButtonClick;
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
