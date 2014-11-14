using System;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using System.Windows.Forms;

namespace Rubberduck.ToDoItems
{
    internal class ToDoItemsMenu : IDisposable
    {
        private readonly VBE _vbe;
        private readonly AddIn _addIn;
        private readonly DockableWindowHost _controlHost;
        private Window _toolWindow;

        private CommandBarButton _todoItemsButton;
        public CommandBarButton ToDoItemsButton { get { return _todoItemsButton; } }

        public ToDoItemsMenu(VBE vbe, AddIn addInInstance)
        {
            _vbe = vbe;
            _addIn = addInInstance;
            _controlHost = new DockableWindowHost();

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
            var control = new ToDoItemsControl(_vbe);
            _toolWindow = CreateToolWindow("ToDo Items", "{9CF1392A-2DC9-48A6-AC0B-E601A9802608}", control);
        }

        private Window CreateToolWindow(string toolWindowCaption, string toolWindowGUID, UserControl toolWindowUserControl)
        {
            //todo: create base class to expose this. Will need to be *protected*.
            Object userControlObject = null;
            DockableWindowHost userControlHost;
            Window toolWindow;
            const string progId = "Rubberduck.DockableWindowHost"; //DockableWindowHost progId

            toolWindow = _vbe.Windows.CreateToolWindow(_addIn, progId, toolWindowCaption, toolWindowGUID, ref userControlObject);

            userControlHost = (DockableWindowHost)userControlObject;
            toolWindow.Visible = true;

            userControlHost.AddUserControl(toolWindowUserControl);

            return toolWindow;

        }

        public void Dispose()
        {
            _controlHost.Dispose();
        }
    }
}
