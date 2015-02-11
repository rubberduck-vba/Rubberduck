using System;
using System.Drawing;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using System.Windows.Forms;
using Rubberduck.Properties;

namespace Rubberduck.UI
{
    [ComVisible(false)]
    public class Menu : IDisposable
    {
        private VBE vbe;
        protected readonly AddIn addInInstance;

        protected VBE IDE { get { return this.vbe; } }

        public Menu(VBE vbe, AddIn addInInstance)
        {
            this.vbe = vbe;
            this.addInInstance = addInInstance;
        }

        protected CommandBarButton AddMenuButton(CommandBarPopup menu, string caption, Bitmap image)
        {
            var result = menu.Controls.Add(MsoControlType.msoControlButton, Temporary: true) as CommandBarButton;
            if (result == null)
            {
                throw new InvalidOperationException("Failed to create menu control.");
            }

            result.Caption = caption;
            SetButtonImage(result, image);

            return result;
        }



        public static void SetButtonImage(CommandBarButton result, Bitmap image)
        {
            result.FaceId = 0;

            if (image != null)
            {
                Clipboard.SetDataObject(image, true);
                result.PasteFace();
            }
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
            for (var i = 1; i <= controls.Count; i++)
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
            _DockableWindowHost userControlHost;
            Window toolWindow;
            const string dockableWindowHostProgId = "Rubberduck.UI.DockableWindowHost"; //DockableWindowHost progId
            const string dockableWindowHostGUID = "9CF1392A-2DC9-48A6-AC0B-E601A9802608";

            toolWindow = this.vbe.Windows.CreateToolWindow(this.addInInstance, dockableWindowHostProgId, toolWindowCaption, dockableWindowHostGUID, ref userControlObject);

            userControlHost = (_DockableWindowHost)userControlObject;
            toolWindow.Visible = true; //window resizing doesn't work without this

            userControlHost.AddUserControl(toolWindowUserControl);

            return toolWindow;

        }

        public void Dispose() { }
    }
}
