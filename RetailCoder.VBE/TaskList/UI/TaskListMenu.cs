using System;
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
            var menuBarControls = this.vbe.CommandBars["Menu Bar"].Controls;
            var toolsMenu = (CommandBarPopup)menuBarControls["Tools"];
            int beforeIndex = FindMenuInsertionIndex(toolsMenu.Controls);
            showTaskListButton = (CommandBarButton)toolsMenu.Controls.Add(Type: MsoControlType.msoControlButton, Before: beforeIndex, Temporary: true);
            showTaskListButton.Caption = "&Task List";

            const int clipboardWithCheck = 837;
            showTaskListButton.FaceId = clipboardWithCheck;

            showTaskListButton.Click += OnShowTaskListButtonClick;

        }

        void OnShowTaskListButtonClick(CommandBarButton ctrl, ref bool CancelDefault)
        {
            TaskListControl taskListControl;

            if ( this.toolWindow == null)
            {
                taskListControl = new TaskListControl(this.vbe);
                toolWindow = CreateToolWindow("Task List", "{9CF1392A-2DC9-48A6-AC0B-E601A9802608}", taskListControl);
            }
            else
            {
                this.toolWindow.Visible = true;
            }
        }

        private Window CreateToolWindow(string toolWindowCaption, string toolWindowGUID, UserControl toolWindowUserControl)
        {
            //todo: create base class to expose this. Will need to be *protected*.
            Object userControlObject = null;
            DockableWindowHost userControlHost;
            Window toolWindow;
            const string progId = "RetailCoderVBE.DockableWindowHost"; //DockableWindowHost progId

            toolWindow = this.vbe.Windows.CreateToolWindow(this.addInInstance, progId, toolWindowCaption, toolWindowGUID, ref userControlObject);

            userControlHost = (DockableWindowHost)userControlObject;
            toolWindow.Visible = true;

            userControlHost.AddUserControl(toolWindowUserControl);

            return toolWindow;

        }

        //todo: need a base menu class that takes in an additonal "before" and/or after param
        private int FindMenuInsertionIndex(CommandBarControls controls)
        {
            for (int i = 1; i <= controls.Count; i++)
            {
                // insert menu before "Window" built-in menu:
                if (controls[i].BuiltIn && controls[i].Caption == "&Macros...")
                {
                    return i;
                }
            }

            return controls.Count;
        }

        public void Dispose()
        {
            this.userControlHost.Dispose();
        }
    }
}
