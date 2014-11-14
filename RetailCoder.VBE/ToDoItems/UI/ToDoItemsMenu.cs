using System;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;

namespace Rubberduck.ToDoItems.UI
{
    internal class ToDoItemsMenu : IDisposable
    {
        private readonly VBE _vbe;
        private readonly AddIn _addIn;
        private readonly DockableWindowHost _controlHost;
        private readonly Window _toolWindow;

        public ToDoItemsMenu(VBE vbe, AddIn addInInstance)
        {
            _vbe = vbe;
            _addIn = addInInstance;
            _controlHost = new DockableWindowHost();

            var control = new ToDoItemsControl(_vbe);
            _toolWindow = CreateToolWindow("ToDo Items", "{9CF1392A-2DC9-48A6-AC0B-E601A9802608}", control);
        }

        public CommandBarButton ToDoItemsButton { get; private set; }

        public void Initialize(CommandBarControls menuControls)
        {
            ToDoItemsButton = menuControls.Add(Type: MsoControlType.msoControlButton, Temporary: true) as CommandBarButton;
            if (ToDoItemsButton == null) return;

            ToDoItemsButton.Caption = "&ToDo Items";
            ToDoItemsButton.BeginGroup = true;

            const int clipboardWithCheck = 837;
            ToDoItemsButton.FaceId = clipboardWithCheck;
            ToDoItemsButton.Click += OnShowTaskListButtonClick;
        }

        void OnShowTaskListButtonClick(CommandBarButton ctrl, ref bool CancelDefault)
        {
            _toolWindow.Visible = true;
        }

        private Window CreateToolWindow(string toolWindowCaption, string toolWindowGuid, UserControl toolWindowUserControl)
        {
            //todo: create base class to expose this. Will need to be *protected*.
            Object userControlObject = null;
            const string progId = "Rubberduck.DockableWindowHost"; //DockableWindowHost progId

            var toolWindow = _vbe.Windows.CreateToolWindow(_addIn, progId, toolWindowCaption, toolWindowGuid, ref userControlObject);
            toolWindow.Visible = true;

            var userControlHost = (DockableWindowHost)userControlObject;
            userControlHost.AddUserControl(toolWindowUserControl);

            return toolWindow;
        }

        public void Dispose()
        {
            _controlHost.Dispose();
        }
    }
}
