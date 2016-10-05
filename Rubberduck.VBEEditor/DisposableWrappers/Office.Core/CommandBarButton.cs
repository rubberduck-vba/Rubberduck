using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.CSharp.RuntimeBinder;

namespace Rubberduck.VBEditor.DisposableWrappers.Office.Core
{
    public class CommandBarButton : CommandBarControl
    {
        public CommandBarButton(Microsoft.Office.Core.CommandBarButton comObject) 
            : base(comObject)
        {
            comObject.Click += comObject_Click;
        }

        public static CommandBarButton FromCommandBarControl(CommandBarControl control)
        {
            return new CommandBarButton((Microsoft.Office.Core.CommandBarButton)control.ComObject);
        }

        private Microsoft.Office.Core.CommandBarButton Button
        {
            get { return (Microsoft.Office.Core.CommandBarButton)ComObject; }
        }

        public event EventHandler<CommandBarButtonClickEventArgs> Click;
        private void comObject_Click(Microsoft.Office.Core.CommandBarButton ctrl, ref bool cancelDefault)
        {
            // todo: confirm whether this fixes the multicast glitch of ParentMenuItemBase.child_Click()
            // "without this hack, handler runs once for each menu item that's hooked up to the command.
            //  hash code is different on every frakkin' click. go figure. I've had it, this is the fix."

            var handler = Click;
            if (handler == null)
            {
                return;
            }

            var args = new CommandBarButtonClickEventArgs(new CommandBarButton(ctrl));
            handler.Invoke(this, args);
            cancelDefault = args.Cancel;
        }

        public bool IsBuiltInFace
        {
            get { return !IsWrappingNullReference && InvokeResult(() => Button.BuiltInFace); }
            set { Invoke(() => Button.BuiltInFace = value); }
        }

        public int FaceId 
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => Button.FaceId); }
            set { Invoke(() => Button.FaceId = value); }
        }

        public string ShortcutText
        {
            get { return IsWrappingNullReference ? string.Empty : InvokeResult(() => Button.ShortcutText); }
            set { Invoke(() => Button.ShortcutText = value); }
        }

        public ButtonState State
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => (ButtonState)Button.State); }
            set { Invoke(() => Button.State = (Microsoft.Office.Core.MsoButtonState)value); }
        }

        public ButtonStyle Style
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => (ButtonStyle)Button.Style); }
            set { Invoke(() => Button.Style = (Microsoft.Office.Core.MsoButtonStyle)value); }
        }

        public Image Picture { get; set; }
        public Image Mask { get; set; }

        public void ApplyIcon()
        {
            Invoke(() =>
            {
                Button.FaceId = 0;
                if (Picture == null || Mask == null)
                {
                    return;
                }

                if (!HasPictureProperty)
                {
                    using (var image = CreateTransparentImage(Picture, Mask))
                    {
                        Clipboard.SetImage(image);
                        Button.PasteFace();
                        Clipboard.Clear();
                    }
                    return;
                }

                Button.Picture = AxHostConverter.ImageToPictureDisp(Picture);
                Button.Mask = AxHostConverter.ImageToPictureDisp(Mask);
            });
        }

        private bool? _hasPictureProperty;
        private bool HasPictureProperty
        {
            get
            {
                if (IsWrappingNullReference)
                {
                    return false;
                }

                if (_hasPictureProperty.HasValue)
                {
                    return _hasPictureProperty.Value;
                }

                try
                {
                    dynamic button = Button;
                    var picture = button.Picture;
                    _hasPictureProperty = true;
                }
                catch (RuntimeBinderException)
                {
                    _hasPictureProperty = false;
                }

                return _hasPictureProperty.Value;
            }
        }

        private static Image CreateTransparentImage(Image image, Image mask)
        {
            //HACK - just blend image with a SystemColors value (mask is ignored)
            //TODO - a real solution would use clipboard formats "Toolbar Button Face" AND "Toolbar Button Mask"
            //because PasteFace apparently needs both to be present on the clipboard
            //However, the clipboard formats are apparently only accessible in English versions of Office
            //https://social.msdn.microsoft.com/Forums/office/en-US/33e97c32-9fc2-4531-b208-67c39ccfb048/question-about-toolbar-button-face-in-pasteface-?forum=vsto

            var output = new Bitmap(image.Width, image.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(output))
            {
                g.Clear(SystemColors.MenuBar);
                g.DrawImage(image, 0, 0);
            }
            return output;
        }
    }
}