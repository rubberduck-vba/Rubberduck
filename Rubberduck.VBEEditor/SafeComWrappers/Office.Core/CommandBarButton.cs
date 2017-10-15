using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.CSharp.RuntimeBinder;
using Rubberduck.VBEditor.SafeComWrappers.MSForms;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;
using ButtonState = Rubberduck.VBEditor.SafeComWrappers.MSForms.ButtonState;

namespace Rubberduck.VBEditor.SafeComWrappers.Office.Core
{
    public class CommandBarButton : CommandBarControl, ICommandBarButton
    {
        public CommandBarButton(Microsoft.Office.Core.CommandBarButton target) 
            : base(target)
        {
        }

        private Microsoft.Office.Core.CommandBarButton Button
        {
            get { return (Microsoft.Office.Core.CommandBarButton)Target; }
        }

        public static ICommandBarButton FromCommandBarControl(ICommandBarControl control)
        {
            return new CommandBarButton((Microsoft.Office.Core.CommandBarButton)control.Target);
        }

        private EventHandler<CommandBarButtonClickEventArgs> _clickHandler; 
        public event EventHandler<CommandBarButtonClickEventArgs> Click
        {
            add
            {
                ((Microsoft.Office.Core.CommandBarButton)Target).Click += Target_Click;
                _clickHandler += value;
                System.Diagnostics.Debug.WriteLine("Added handler for: {0} '{1}' (tag: {2}, hashcode:{3})", Parent.Name, Target.Caption, Tag, Target.GetHashCode());
            }
            remove
            {
                try
                {
                    ((Microsoft.Office.Core.CommandBarButton)Target).Click -= Target_Click;
                }
                catch
                {
                    // he's gone, dave.
                }
                _clickHandler -= value;
                System.Diagnostics.Debug.WriteLine("Removed handler for: {0} '{1}' (tag: {2}, hashcode:{3})", Parent.GetType().Name, Target.Caption, Tag, Target.GetHashCode());
            }
        }

        private void Target_Click(Microsoft.Office.Core.CommandBarButton ctrl, ref bool cancelDefault)
        {
            var handler = _clickHandler;
            if (handler == null || IsWrappingNullReference)
            {
                return;
            }

            System.Diagnostics.Debug.Assert(handler.GetInvocationList().Length == 1, "Multicast delegate is registered more than once.");

            //note: event is fired for every parent the command exists under. not sure why.
            System.Diagnostics.Debug.WriteLine("Executing handler for: {0} '{1}' (tag: {2}, hashcode:{3})", Parent.GetType().Name, Target.Caption, Tag, Target.GetHashCode());

            var button = new CommandBarButton(ctrl);
            var args = new CommandBarButtonClickEventArgs(button);
            handler.Invoke(this, args);
            cancelDefault = args.Cancel;
            //button.Release(final:true);
        }

        public bool IsBuiltInFace
        {
            get { return !IsWrappingNullReference && Button.BuiltInFace; }
            set { if (!IsWrappingNullReference) Button.BuiltInFace = value; }
        }

        public int FaceId 
        {
            get { return IsWrappingNullReference ? 0 : Button.FaceId; }
            set { if (!IsWrappingNullReference) Button.FaceId = value; }
        }

        public string ShortcutText
        {
            get { return IsWrappingNullReference ? string.Empty : Button.ShortcutText; }
            set { if (!IsWrappingNullReference) Button.ShortcutText = value; }
        }

        public ButtonState State
        {
            get { return IsWrappingNullReference ? 0 : (ButtonState)Button.State; }
            set { if (!IsWrappingNullReference) Button.State = (Microsoft.Office.Core.MsoButtonState)value; }
        }

        public ButtonStyle Style
        {
            get { return IsWrappingNullReference ? 0 : (ButtonStyle)Button.Style; }
            set { if (!IsWrappingNullReference) Button.Style = (Microsoft.Office.Core.MsoButtonStyle)value; }
        }

        public Image Picture { get; set; }
        public Image Mask { get; set; }

        public void ApplyIcon()
        {
            if (IsWrappingNullReference) return;
            
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