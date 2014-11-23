using System;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using System.Windows.Forms;
using System.Collections.Generic;

namespace Rubberduck.UI
{
    internal class ToDoItemsMenu 
    {
        private readonly VBE _vbe;
        private readonly AddIn _addIn;
        private Window _toolWindow;
        private Config.ToDoListSettings _settings;

        private CommandBarButton _todoItemsButton;
        public CommandBarButton ToDoItemsButton { get { return _todoItemsButton; } }

        public ToDoItemsMenu(VBE vbe, AddIn addInInstance, Config.ToDoListSettings settings)
        {
            _vbe = vbe;
            _addIn = addInInstance;
            _settings = settings;
        }

        public void Initialize(CommandBarControls menuControls)
        {
            _todoItemsButton = menuControls.Add(Type: MsoControlType.msoControlButton, Temporary: true) as CommandBarButton;
            _todoItemsButton.Caption = "&ToDo Items";
            _todoItemsButton.BeginGroup = true;

            const int clipboardWithCheck = 837;
            _todoItemsButton.FaceId = clipboardWithCheck;
            _todoItemsButton.Click += OnShowTaskListButtonClick;
        }

        void OnShowTaskListButtonClick(CommandBarButton ctrl, ref bool CancelDefault)
        {
            if (_toolWindow == null)
            {
                InitializeWindow();
            }

            _toolWindow.Visible = true;
        }

        private void InitializeWindow()
        {
            var markers = new List<Config.ToDoMarker>(_settings.ToDoMarkers);
            var control = new ToDoItemsControl(_vbe, markers);
            _toolWindow = CreateToolWindow("ToDo Items", control);
        }

        private Window CreateToolWindow(string toolWindowCaption, UserControl toolWindowUserControl)
        {
            //todo: create base class to expose this. Will need to be *protected*.
            Object userControlObject = null;
            DockableWindowHost userControlHost;
            Window toolWindow;
            const string progId = "Rubberduck.UI.DockableWindowHost";
            const string dockableHostGuid = "{9CF1392A-2DC9-48A6-AC0B-E601A9802608}";

            toolWindow = _vbe.Windows.CreateToolWindow(_addIn, progId, toolWindowCaption, dockableHostGuid, ref userControlObject);

            userControlHost = (DockableWindowHost)userControlObject;
            toolWindow.Visible = true; //window resizing doesn't work without this

            userControlHost.AddUserControl(toolWindowUserControl);

            return toolWindow;

        }
    }
}
