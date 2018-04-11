using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;
using MSO = Microsoft.Office.Core;
using NLog;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.MSForms;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;
using ButtonState = Rubberduck.VBEditor.SafeComWrappers.MSForms.ButtonState;

namespace Rubberduck.VBEditor.SafeComWrappers.Office.Core
{
    public class CommandBarButton : SafeEventedComWrapper<MSO.CommandBarButton, MSO.ICommandBarsEvents>, ICommandBarButton, MSO.ICommandBarButtonEvents
    {
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();
        public const bool AddCommandBarControlsTemporarily = false;

        public CommandBarButton(MSO.CommandBarButton target, bool rewrapping = false) 
            : base(target, rewrapping)
        {
        }
        
        private MSO.CommandBarButton Button => Target;

        public static CommandBarButton FromCommandBarControl(ICommandBarControl control)
        {
            return new CommandBarButton((MSO.CommandBarButton)control.Target, rewrapping: true);
        }
        
        public bool IsBuiltInFace
        {
            get => !IsWrappingNullReference && Button.BuiltInFace;
            set
            {
                if (!IsWrappingNullReference)
                {
                    Button.BuiltInFace = value;
                }
            }
        }

        public int FaceId 
        {
            get => IsWrappingNullReference ? 0 : Button.FaceId;
            set
            {
                if (!IsWrappingNullReference)
                {
                    Button.FaceId = value;
                }
            }
        }

        public string ShortcutText
        {
            get => IsWrappingNullReference ? string.Empty : Button.ShortcutText;
            set
            {
                if (!IsWrappingNullReference)
                {
                    Button.ShortcutText = value;
                }
            }
        }

        public ButtonState State
        {
            get => IsWrappingNullReference ? 0 : (ButtonState)Button.State;
            set
            {
                if (!IsWrappingNullReference)
                {
                    Button.State = (MSO.MsoButtonState)value;
                }
            }
        }

        public ButtonStyle Style
        {
            get => IsWrappingNullReference ? 0 : (ButtonStyle)Button.Style;
            set
            {
                if (!IsWrappingNullReference)
                {
                    Button.Style = (MSO.MsoButtonStyle)value;
                }
            }
        }

        public Image Picture { get; set; }
        public Image Mask { get; set; }

        public void ApplyIcon()
        {
            if (IsWrappingNullReference)
            {
                return;
            }

            Button.FaceId = 0;
            if (Picture == null || Mask == null)
            {
                return;
            }

            if (!HasPictureProperty)
            {
                using (var image = CreateTransparentImage(Picture))
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

                catch (COMException)
                {
                    _hasPictureProperty = false;
                }

                return _hasPictureProperty.Value;
            }
        }

        private static Image CreateTransparentImage(Image image)
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

        public bool BeginsGroup
        {
            get => !IsWrappingNullReference && Target.BeginGroup;
            set
            {
                if (!IsWrappingNullReference)
                {
                    Target.BeginGroup = value;
                }
            }
        }

        public bool IsBuiltIn => !IsWrappingNullReference && Target.BuiltIn;

        public string Caption
        {
            get => IsWrappingNullReference ? string.Empty : Target.Caption;
            set
            {
                if (!IsWrappingNullReference)
                {
                    Target.Caption = CommandBarControlCaptionGuard.ApplyGuard(value);
                }
            }
        }

        public string DescriptionText
        {
            get => IsWrappingNullReference ? string.Empty : Target.DescriptionText;
            set
            {
                if (!IsWrappingNullReference)
                {
                    Target.DescriptionText = value;
                }
            }
        }

        public bool IsEnabled
        {
            get => !IsWrappingNullReference && Target.Enabled;
            set
            {
                if (!IsWrappingNullReference)
                {
                    Target.Enabled = value;
                }
            }
        }

        public int Height
        {
            get => Target.Height;
            set
            {
                if (!IsWrappingNullReference)
                {
                    Target.Height = value;
                }
            }
        }

        public int Id => IsWrappingNullReference ? 0 : Target.Id;

        public int Index => IsWrappingNullReference ? 0 : Target.Index;

        public int Left => IsWrappingNullReference ? 0 : Target.Left;

        public string OnAction
        {
            get => IsWrappingNullReference ? string.Empty : Target.OnAction;
            set
            {
                if (!IsWrappingNullReference)
                {
                    Target.OnAction = value;
                }
            }
        }

        public ICommandBar Parent => new CommandBar(IsWrappingNullReference ? null : Target.Parent);

        public string Parameter
        {
            get => IsWrappingNullReference ? string.Empty : Target.Parameter;
            set
            {
                if (!IsWrappingNullReference)
                {
                    Target.Parameter = value;
                }
            }
        }

        public int Priority
        {
            get => IsWrappingNullReference ? 0 : Target.Priority;
            set
            {
                if (!IsWrappingNullReference)
                {
                    Target.Priority = value;
                }
            }
        }

        public string Tag
        {
            get => Target?.Tag;
            set
            {
                if (!IsWrappingNullReference)
                {
                    Target.Tag = value;
                }
            }
        }

        public string TooltipText
        {
            get => IsWrappingNullReference ? string.Empty : Target.TooltipText;
            set
            {
                if (!IsWrappingNullReference)
                {
                    Target.TooltipText = value;
                }
            }
        }

        public int Top => IsWrappingNullReference ? 0 : Target.Top;

        public ControlType Type => IsWrappingNullReference ? 0 : (ControlType)Target.Type;

        public bool IsVisible
        {
            get => !IsWrappingNullReference && Target.Visible;
            set
            {
                if (!IsWrappingNullReference)
                {
                    Target.Visible = value;
                }
            }
        }

        public int Width
        {
            get => IsWrappingNullReference ? 0 : Target.Width;
            set
            {
                if (!IsWrappingNullReference)
                {
                    Target.Width = value;
                }
            }
        }

        public bool IsPriorityDropped => (!IsWrappingNullReference) && Target.IsPriorityDropped;

        public void Delete()
        {
            if (!IsWrappingNullReference)
            {
                DetachEvents();
                Target.Delete(AddCommandBarControlsTemporarily);
            }
        }

        public void Execute()
        {
            if (!IsWrappingNullReference)
            {
                Target.Execute();
            }
        }
        
        public bool Equals(ICommandBarControl other)
        {
            return Equals(other as SafeComWrapper<MSO.CommandBarControl>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : HashCode.Compute(Type, Id, Index, IsBuiltIn, Target.Parent);
        }

        public override bool Equals(ISafeComWrapper<MSO.CommandBarButton> other)
        {
            return IsEqualIfNull(other) ||
                   (other != null
                    && (int)other.Target.Type == (int)Type
                    && other.Target.Id == Id
                    && other.Target.Index == Index
                    && other.Target.BuiltIn == IsBuiltIn
                    && ReferenceEquals(other.Target.Parent, Target.Parent));
        }

        private object _eventLock = new object();
        private event EventHandler<CommandBarButtonClickEventArgs> _click;
        public event EventHandler<CommandBarButtonClickEventArgs> Click
        {
            add
            {
                lock (_eventLock)
                {
                    if (_click != null && _click.GetInvocationList().Length == 0)
                    {
                        AttachEvents();
                    }
                    _click += value;
                }
            }
            remove
            {
                lock (_eventLock)
                {
                    if (_click != null && _click.GetInvocationList().Length == 0)
                    {
                        DetachEvents();
                    }
                    _click -= value;
                }
            }
        }

        void MSO.ICommandBarButtonEvents.Click(MSO.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            var button = new CommandBarButton(Ctrl);
            var handler = _click;
            if (handler == null || IsWrappingNullReference)
            {
                button.Dispose();
                return;
            }

            System.Diagnostics.Debug.Assert(handler.GetInvocationList().Length == 1, "Multicast delegate is registered more than once.");

            var args = new CommandBarButtonClickEventArgs(button);
            handler.Invoke(this, args);
            CancelDefault = args.Cancel;
        }

        public event EventHandler Disposing;
        protected override void Dispose(bool disposing)
        {
            Disposing?.Invoke(this, EventArgs.Empty);
            base.Dispose(disposing);
        }
    }
}