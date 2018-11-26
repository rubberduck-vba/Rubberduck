extern alias Office_v8;
using System;
using System.Drawing;
using NLog;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using MSO = Office_v8::Office;
using VB = Microsoft.Vbe.Interop.VB6;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.Office8
{
    public sealed class CommandBarButton : SafeRedirectedEventedComWrapper<MSO.CommandBarButton, VB.CommandBarEvents, VB._dispCommandBarControlEvents>, ICommandBarButton, VB._dispCommandBarControlEvents
    {
        private readonly CommandBarControl _control;
        private readonly IVBE _vbe;

        public const bool AddCommandBarControlsTemporarily = false;        

        // Command bar click event is sourced from VBE.Events.CommandBarEvents[index]
        // where index is the command bar button COM object.
        public CommandBarButton(MSO.CommandBarButton target, IVBE vbe, bool rewrapping = false) 
            : base(target, rewrapping)
        {
            _control = new CommandBarControl(target, vbe, true);
            _vbe = vbe;
        }
        
        private MSO.CommandBarButton Button => Target;
        
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

            // TODO - find a solution that doesn't rely on .Picture (unavailable) or .PasteFace (no-op, exception-prone and destructive)
            // For now, menu and commandbar icons will be absent in VB6.
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
            get => _control.BeginsGroup;
            set => _control.BeginsGroup = value;
        }

        public bool IsBuiltIn => _control.IsBuiltIn;

        public string Caption
        {
            get => _control.Caption;
            set => _control.Caption = value;
        }

        public string DescriptionText
        {
            get => _control.DescriptionText;
            set => _control.DescriptionText=value;
        }

        public bool IsEnabled
        {
            get => _control.IsEnabled;
            set=> _control.IsEnabled = value;
        }

        public int Height
        {
            get => _control.Height;
            set => _control.Height = value;
        }

        public int Id => _control.Id;

        public int Index => _control.Index;

        public int Left => _control.Left;

        public string OnAction
        {
            get => _control.OnAction;
            set => _control.OnAction = value;
        }

        public ICommandBar Parent => _control.Parent;

        public string Parameter
        {
            get => _control.Parameter;
            set => _control.Parameter = value;
        }

        public int Priority
        {
            get => _control.Priority;
            set => _control.Priority = value;
        }

        public string Tag
        {
            get => _control.Tag;
            set => _control.Tag = value;
        }

        public string TooltipText
        {
            get => _control.TooltipText;
            set => _control.TooltipText = value;
        }

        public int Top => _control.Top;

        public ControlType Type => _control.Type;

        public bool IsVisible
        {
            get => _control.IsVisible;
            set => _control.IsVisible = value;
        }

        public int Width
        {
            get => _control.Width;
            set => _control.Width = value;
        }

        public void Delete()
        {
            if (!IsWrappingNullReference)
            {
                DetachEvents();
            }
            _control.Delete();
        }

        public void Execute()
        {
            _control.Execute();
        }
        
        public bool Equals(ICommandBarControl other)
        {
            return _control.Equals(other);
        }

        public override bool Equals(ISafeComWrapper<MSO.CommandBarButton> other)
        {
            return _control.Equals(other);
        }
        
        public override int GetHashCode()
        {
            return _control.GetHashCode();
        }

        private readonly object _eventLock = new object();
        private event EventHandler<CommandBarButtonClickEventArgs> _click;
        public event EventHandler<CommandBarButtonClickEventArgs> Click
        {
            add
            {
                lock (_eventLock)
                {
                    _click += value;
                    if (_click != null && _click.GetInvocationList().Length == 1)
                    {
                        // First subscriber attached - attach COM events
                        AttachEvents();
                    }
                }
            }
            remove
            {
                lock (_eventLock)
                {
                    _click -= value;
                    if (_click == null || _click.GetInvocationList().Length == 0)
                    {
                        // Last subscriber detached - detach COM events 
                        DetachEvents();
                    }
                }
            }
        }

        // Event is routed here from VBE.Events.CommandBarEvents[] - see ctor comment.
        void VB._dispCommandBarControlEvents.Click(object Ctrl, ref bool Handled, ref bool CancelDefault)
        {            
            var handler = _click;
            if (handler == null || IsWrappingNullReference)
            {                
                return;
            }

            using (var button = new CommandBarButton((MSO.CommandBarButton) Ctrl, _vbe))
            {
                System.Diagnostics.Debug.Assert(handler.GetInvocationList().Length == 1,
                    "Multicast delegate is registered more than once.");

                var args = new CommandBarButtonClickEventArgs(button);
                handler.Invoke(this, args);
                CancelDefault = args.Cancel;
                Handled = args.Handled;
            }
        }
		
        public event EventHandler Disposing;
        protected override void Dispose(bool disposing)
        {
            Disposing?.Invoke(this, EventArgs.Empty);
            base.Dispose(disposing);
            _control.Dispose();
        }

        public override void AttachEvents()
        {
            // Cast to VB6 VBE SafeComWrapper as  events are not exposed on IVBE as they are only safe to use in VB6
            using (var events = ((VB6.VBE)_vbe).Events)
            using (var commandBarEvents = events.CommandBarEvents)
            {
                // Disposal of buttonEvents is handled by the base class
                var buttonEvents = commandBarEvents[Target] as IEventSource<VB.CommandBarEvents>;
                AttachEvents(buttonEvents, this);
            }
        }
    }
}