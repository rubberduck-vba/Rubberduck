using System;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using System.Windows.Forms;

namespace Rubberduck.ToDoItems
{
    internal class ToDoItemsMenu : Rubberduck.Menu
    {
        private Window toolWindow;

        public ToDoItemsMenu(VBE vbe, AddIn addInInstance):base(vbe, addInInstance){}

        private CommandBarButton showTaskListButton;
        public CommandBarButton ShowTaskListButton { get { return this.showTaskListButton; } }

        public void Initialize(CommandBarControls menuControls)
        {
            var menuBarControls = this.IDE.CommandBars["Menu Bar"].Controls;
            var toolsMenu = (CommandBarPopup)menuBarControls["Tools"];
            int beforeIndex = FindMenuInsertionIndex(toolsMenu.Controls, "&Macros...");
            
            showTaskListButton = AddMenuButton(toolsMenu);
            showTaskListButton.Caption = "&Task List";

            const int clipboardWithCheck = 837;
            showTaskListButton.FaceId = clipboardWithCheck;

            showTaskListButton.Click += OnShowTaskListButtonClick;

        }

        void OnShowTaskListButtonClick(CommandBarButton ctrl, ref bool CancelDefault)
        {
            ToDoItemsControl taskListControl;

            if ( this.toolWindow == null)
            {
                taskListControl = new ToDoItemsControl(this.IDE);
                toolWindow = CreateToolWindow("Task List",  taskListControl);
            }                                             
            else
            {
                this.toolWindow.Visible = true;
            }
        }
    }
}
