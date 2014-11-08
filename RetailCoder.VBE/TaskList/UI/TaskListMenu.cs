using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using System.Windows.Forms;

namespace RetailCoderVBE.TaskList
{
    internal class TaskListMenu : IDisposable
    {
        private VBE vbe;
        private DockableWindowHost userControlHost;
        private AddIn addInInstance;
        private Window toolWindow;

        public TaskListMenu(VBE vbe, AddIn addInInstance)
        {
            this.vbe = vbe;
            this.addInInstance = addInInstance;
            this.userControlHost = new DockableWindowHost();
        }

        private CommandBarButton showTaskListButton;
        public CommandBarButton ShowTaskListButton { get { return this.showTaskListButton; } }

        public void Initialize()
        {
            //todo: insert menu item after CallStack item
            //todo: provide icon
            var menuBarControls = this.vbe.CommandBars["Menu Bar"].Controls;
            CommandBarPopup viewMenu = (CommandBarPopup)menuBarControls["View"];
            showTaskListButton = (CommandBarButton)viewMenu.Controls.Add(Type: MsoControlType.msoControlButton, Temporary: true);
            showTaskListButton.Caption = "&Task List";

            showTaskListButton.Click += OnShowTaskListButtonClick;

        }

        void OnShowTaskListButtonClick(CommandBarButton ctrl, ref bool CancelDefault)
        {
            TaskListControl taskListControl;

            try
            {
                if ( this.toolWindow == null)
                {
                    taskListControl = new TaskListControl(this.vbe);
                    //todo: implement tasklist
                    toolWindow = CreateToolWindow("My dockable window", "{9CF1392A-2DC9-48A6-AC0B-E601A9802608}", taskListControl);
                }
                else
                {
                    this.toolWindow.Visible = true;
                }

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private Window CreateToolWindow(string toolWindowCaption, string toolWindowGUID, UserControl toolWindowUserControl)
        {
            //todo: create base class to expose this. Will need to be *protected*.
            Object userControlObject = null;
            DockableWindowHost userControlHost;
            Window toolWindow;
            string progId = "RetailCoderVBE.DockableWindowHost"; //DockableWindowHost progId

            toolWindow = this.vbe.Windows.CreateToolWindow(this.addInInstance, progId, toolWindowCaption, toolWindowGUID, ref userControlObject);

            userControlHost = (DockableWindowHost)userControlObject;
            toolWindow.Visible = true;

            userControlHost.AddUserControl(toolWindowUserControl);

            return toolWindow;

        }

        public void Dispose()
        {
            this.userControlHost.Dispose();
        }
    }
}
