using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using CommandBarButtonClickEvent = Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler;

namespace Rubberduck.UI
{
    public class Menu : IDisposable
    {
        private readonly VBE _vbe;
        protected readonly AddIn AddIn;

        protected VBE IDE { get { return this._vbe; } }

        protected Menu(VBE vbe, AddIn addIn)
        {
            AddIn = addIn;
            _vbe = vbe;
        }

        private CommandBarButton AddButton(CommandBarPopup parentMenu, string caption)
        {
            var button = parentMenu.Controls.Add(MsoControlType.msoControlButton, Temporary: true) as CommandBarButton;
            button.Caption = caption;

            return button;
        }

        protected CommandBarButton AddButton(CommandBarPopup parentMenu, string caption, bool beginGroup, CommandBarButtonClickEvent buttonClickHandler)
        {
            var button = AddButton(parentMenu, caption);
            button.BeginGroup = beginGroup;
            button.Click += buttonClickHandler;

            return button;
        }

        protected CommandBarButton AddButton(CommandBarPopup parentMenu, string caption, bool beginGroup, CommandBarButtonClickEvent buttonClickHandler, int faceId)
        {
            var button = AddButton(parentMenu, caption, beginGroup, buttonClickHandler);
            button.FaceId = faceId;

            return button;
        }

        protected CommandBarButton AddButton(CommandBarPopup parentMenu, string caption, bool beginGroup, CommandBarButtonClickEvent buttonClickHandler, Bitmap image)
        {
            var button = AddButton(parentMenu, caption, beginGroup, buttonClickHandler);
            SetButtonImage(button, image);

            return button;
        }

        public static void SetButtonImage(CommandBarButton button, Bitmap image)
        {
            button.FaceId = 0;

            if (image != null)
            {
                Clipboard.SetDataObject(image, true);
                button.PasteFace();
            }
        }

        /// <summary>
        /// Finds the index for insertion in a given CommandBarControls collection.
        /// Returns the last position if the given beforeControl caption is not found.
        /// </summary>
        /// <param name="controls">The collection to insert into.</param>
        /// <param name="beforeId">Caption of the control to insert before.</param>
        /// <returns></returns>
        protected int FindMenuInsertionIndex(CommandBarControls controls, int beforeId)
        {
            for (var i = 1; i <= controls.Count; i++)
            {
                if (controls[i].BuiltIn && controls[i].Id == beforeId)
                {
                    return i;
                }
            }

            return controls.Count;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {            
        }
    }
}
