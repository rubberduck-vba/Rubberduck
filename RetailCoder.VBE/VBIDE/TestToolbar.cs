using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Vbe.Interop;
using Microsoft.Office.Core;

namespace RetailCoderVBE.VBIDE
{
    internal class TestToolbar
    {
        private CommandBarButton _runAllTestsButton;
        public CommandBarButton RunAllTestsButton { get { return _runAllTestsButton; } }

        private CommandBarButton _windowsTestExplorerButton;
        public CommandBarButton WindowsTestExplorerButton { get { return _windowsTestExplorerButton; } }
        
        public void Initialize(VBE vbe)
        {
            var commandBar = vbe.CommandBars.Add(Name: "Test", Temporary: true);

            _windowsTestExplorerButton = AddToolbarButton(commandBar);
            _windowsTestExplorerButton.TooltipText = "Test Explorer";
            _windowsTestExplorerButton.BeginGroup = true;
            _windowsTestExplorerButton.FaceId = 3170; // 305; // a "document" icon, with a green checkmark and a red cross
            _windowsTestExplorerButton.Click += OnTestExplorerButtonClick;

            _runAllTestsButton = AddToolbarButton(commandBar);
            _runAllTestsButton.Caption = "Run all tests";
            _runAllTestsButton.FaceId = 186; // a "play" icon
            _runAllTestsButton.Click += OnRunAllTestsButtonClick;

            commandBar.Visible = true;
        }

        private CommandBarButton AddToolbarButton(CommandBar commandBar)
        {
            return commandBar.Controls.Add(Type: MsoControlType.msoControlButton) as CommandBarButton;
        }

        private void OnButtonClick(EventHandler clickEvent)
        {
            var handler = clickEvent;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }

        public event EventHandler OnTestExporer;
        void OnTestExplorerButtonClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            OnButtonClick(OnTestExporer);
        }

        public event EventHandler OnRunAllTests;
        void OnRunAllTestsButtonClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            OnButtonClick(OnRunAllTests);
        }
    }
}
