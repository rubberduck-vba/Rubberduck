using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using System.Windows.Forms;

namespace Rubberduck
{
    class Menu : IDisposable
    {
        private VBE vbe;
        protected AddIn addInInstance;

        protected VBE IDE { get { return this.vbe; } }

        public Menu(VBE vbe, AddIn addInInstance)
        {
            this.vbe = vbe;
            this.addInInstance = addInInstance;
        }

        protected CommandBarButton AddMenuButton(CommandBarPopup menu)
        {
            return menu.Controls.Add(Type: MsoControlType.msoControlButton, Temporary: true) as CommandBarButton;
        }

        /// <summary>
        /// Finds the index for insertion in a given CommandBarControls collection.
        /// Returns the last position if the given beforeControl caption is not found.
        /// </summary>
        /// <param name="controls">The collection to insert into.</param>
        /// <param name="beforeControl">Caption of the control to insert before.</param>
        /// <returns></returns>
        protected int FindMenuInsertionIndex(CommandBarControls controls, string beforeControl)
        {
            for (int i = 1; i <= controls.Count; i++)
            {
                if (controls[i].BuiltIn && controls[i].Caption == beforeControl)
                {
                    return i;
                }
            }

            return controls.Count;
        }

        /// <summary>
        /// Attaches a user control to a native window through a new DockableWindowHost.
        /// </summary>
        /// <param name="toolWindowCaption">Text to display as the window title.</param>
        /// <param name="toolWindowUserControl">User control to attach to the window.</param>
        /// <returns>Microsoft.Vbe.Interop.Window</returns>
        protected Window CreateToolWindow(string toolWindowCaption, UserControl toolWindowUserControl)
        {
            Object userControlObject = null;
            DockableWindowHost userControlHost;
            Window toolWindow;
            const string dockableWindowHostProgId = "Rubberduck.DockableWindowHost"; //DockableWindowHost progId
            const string dockableWindowHostGUID = "9CF1392A-2DC9-48A6-AC0B-E601A9802608";

            toolWindow = this.vbe.Windows.CreateToolWindow(this.addInInstance, dockableWindowHostProgId, toolWindowCaption, dockableWindowHostGUID, ref userControlObject);

            userControlHost = (DockableWindowHost)userControlObject;
            toolWindow.Visible = true; //window resizing doesn't work without this

            userControlHost.AddUserControl(toolWindowUserControl);

            return toolWindow;

        }

        public void Dispose() { }
    }
}
