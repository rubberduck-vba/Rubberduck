using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;

namespace RetailCoderVBE.VBIDE
{
    [ComVisible(false)]
    public class TestMenu
    {
        // 2743: play icon with stopwatch
        // 3039: module icon || 3119 || 621 || 589 || 472
        // 3170: class module icon

        private CommandBarButton _runAllTestsButton;
        public CommandBarButton RunAllTestsButton { get { return _runAllTestsButton; } }

        private CommandBarButton _windowsTestExplorerButton;
        public CommandBarButton WindowsTestExplorerButton { get { return _windowsTestExplorerButton; } }

        public void Initialize(VBE vbe)
        {
            var menuBarControls = vbe.CommandBars[1].Controls;
            var beforeIndex = FindMenuInsertionIndex(menuBarControls);
            var menu = menuBarControls.Add(Type: MsoControlType.msoControlPopup, Before: beforeIndex, Temporary: true) as CommandBarPopup;
            menu.Caption = "Te&st";

            _windowsTestExplorerButton = AddMenuButton(menu);
            _windowsTestExplorerButton.Caption = "&Test Explorer";
            _windowsTestExplorerButton.FaceId = 3170; // 305; // a "document" icon, with a green checkmark and a red cross
            _windowsTestExplorerButton.Click += OnTestExplorerButtonClick;

            _runAllTestsButton = AddMenuButton(menu);
            _runAllTestsButton.BeginGroup = true;
            _runAllTestsButton.Caption = "&Run All Tests";
            _runAllTestsButton.FaceId = 186; // a "play" icon
            _runAllTestsButton.Click += OnRunAllTestsButtonClick;
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

        private CommandBarButton AddMenuButton(CommandBarPopup menu)
        {
            return menu.Controls.Add(Type: MsoControlType.msoControlButton) as CommandBarButton;
        }

        private void OnButtonClick(EventHandler clickEvent)
        {
            var handler = clickEvent;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }

        public event EventHandler OnNewTestClass;
        void OnNewTestModuleButtonClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            OnButtonClick(OnNewTestClass);
        }

        public event EventHandler OnRunSelectedTests;
        void OnRunSelectedTestsButtonClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            OnButtonClick(OnRunSelectedTests);
        }

        public event EventHandler OnRunAllTests;
        void OnRunAllTestsButtonClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            OnButtonClick(OnRunAllTests);
        }

        public event EventHandler OnRunFailedTests;
        void OnRunFailedTestsButtonClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            OnButtonClick(OnRunFailedTests);
        }

        public event EventHandler OnRunNotRunTests;
        void OnRunNotRunTestsButtonClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            OnButtonClick(OnRunNotRunTests);
        }

        public event EventHandler OnRunPassedTests;
        void OnRunPassedTestsButtonClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            OnButtonClick(OnRunPassedTests);
        }

        public event EventHandler OnRepeatLastRun;
        void OnRepeatLastRunButtonClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            OnButtonClick(OnRepeatLastRun);
        }

        public event EventHandler OnTestExporer;
        void OnTestExplorerButtonClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            OnButtonClick(OnTestExporer);
        }
    }
}
