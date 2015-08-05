using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Core;
using stdole;

namespace Rubberduck.UI.Commands
{
    public class RubberduckMenuCommand : IRubberduckMenuCommand
    {
        private readonly IList<CommandBarButton> _buttons;

        /// <summary>
        /// Creates a new menu command.
        /// </summary>
        public RubberduckMenuCommand()
        {
            _buttons = new List<CommandBarButton>();
        }

        public void AddCommandBarButton(CommandBarControls parent, string caption, bool beginGroup = false, int beforeIndex = -1, Image image = null, Image mask = null)
        {
            if (image != null && mask == null)
            {
                throw new ArgumentNullException("'image' cannot be null if 'mask' is non-null.");
            }
            if (image == null && mask != null)
            {
                throw new ArgumentNullException("'mask' cannot be null if 'image' is non-null.");
            }

            var button = (CommandBarButton) (beforeIndex == -1
                ? parent.Add(MsoControlType.msoControlButton, Temporary: true)
                : parent.Add(MsoControlType.msoControlButton, Before: beforeIndex, Temporary: true));

            button.BeginGroup = beginGroup;
            button.Caption = caption;

            if (image != null)
            {
                SetButtonImage(button, image, mask);
            }

            button.Click += button_Click;
            _buttons.Add(button);
        }

        public void Release()
        {
            foreach (var button in _buttons)
            {
                button.Click -= button_Click;
                try
                {
                    button.Delete();
                    Marshal.ReleaseComObject(button);
                }
                catch (COMException)
                {
                    // just let it be
                }
            }

            _buttons.Clear();
        }

        private void button_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            OnRequestExecute();
        }

        public event EventHandler RequestExecute;
        public void OnRequestExecute()
        {
            var handler = RequestExecute;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }

        private void SetButtonImage(CommandBarButton button, Image image, Image mask)
        {
            button.FaceId = 0;
            button.Picture = AxHostConverter.ImageToPictureDisp(image);
            button.Mask = AxHostConverter.ImageToPictureDisp(mask);
        }

        private class AxHostConverter : AxHost
        {
            private AxHostConverter() : base("") { }

            static public IPictureDisp ImageToPictureDisp(Image image)
            {
                return (IPictureDisp)GetIPictureDispFromPicture(image);
            }

            static public Image PictureDispToImage(IPictureDisp pictureDisp)
            {
                return GetPictureFromIPicture(pictureDisp);
            }
        }
    }
}