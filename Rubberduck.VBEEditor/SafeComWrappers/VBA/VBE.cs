using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.Native;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;
using IAddIns = Rubberduck.VBEditor.SafeComWrappers.Abstract.IAddIns;
using IWindow = Rubberduck.VBEditor.SafeComWrappers.Abstract.IWindow;
using IWindows = Rubberduck.VBEditor.SafeComWrappers.Abstract.IWindows;
using VB = Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class VBE : SafeComWrapper<VB.VBE>, IVBE
    {
        // ReSharper disable once PrivateFieldCanBeConvertedToLocalVariable
        private readonly WinEvents.WinEventDelegate _events;
        private readonly IntPtr _hook = IntPtr.Zero;
        private readonly IntPtr _hwnd = IntPtr.Zero;

        public VBE(VB.VBE target)
            :base(target)
        {
            _events = WinEventProc;
            _hwnd = new IntPtr(target.MainWindow.HWnd);
            _hook = WinEvents.SetWinEventHook((uint)WinEvents.EventConstant.EVENT_MIN,
                (uint)WinEvents.EventConstant.EVENT_MAX, IntPtr.Zero, Marshal.GetFunctionPointerForDelegate(_events), 0, 0,
                (uint)WinEvents.WinEventFlags.WINEVENT_OUTOFCONTEXT);
        }

        public string Version
        {
            get { return IsWrappingNullReference ? string.Empty : Target.Version; }
        }

        public ICodePane ActiveCodePane
        {
            get { return new CodePane(IsWrappingNullReference ? null : Target.ActiveCodePane); }
            set { Target.ActiveCodePane = (VB.CodePane)value.Target; }
        }

        public IVBProject ActiveVBProject
        {
            get { return new VBProject(IsWrappingNullReference ? null : Target.ActiveVBProject); }
            set { Target.ActiveVBProject = (VB.VBProject) value.Target; }
        }

        public IWindow ActiveWindow
        {
            get { return new Window(IsWrappingNullReference ? null : Target.ActiveWindow); }
        }

        public IAddIns AddIns
        {
            get { return new AddIns(IsWrappingNullReference ? null : Target.Addins); }
        }

        public ICodePanes CodePanes
        {
            get { return new CodePanes(IsWrappingNullReference ? null : Target.CodePanes); }
        }

        public ICommandBars CommandBars
        {
            get { return new CommandBars(IsWrappingNullReference ? null : Target.CommandBars); }
        }

        public IWindow MainWindow
        {
            get
            {
                try
                {
                    return new Window(IsWrappingNullReference ? null : Target.MainWindow);
                }
                catch (InvalidComObjectException)
                {
                    return null;
                }
            }
        }

        public IVBComponent SelectedVBComponent
        {
            get { return new VBComponent(IsWrappingNullReference ? null : Target.SelectedVBComponent); }
        }

        public IVBProjects VBProjects
        {
            get { return new VBProjects(IsWrappingNullReference ? null : Target.VBProjects); }
        }

        public IWindows Windows
        {
            get { return new Windows(IsWrappingNullReference ? null : Target.Windows); }
        }

        public Guid EventsInterfaceId { get { throw new NotImplementedException(); } }

        public override void Release(bool final = false)
        {
            WinEvents.UnhookWinEvent(_hook);
            if (!IsWrappingNullReference)
            {
                VBProjects.Release();
                CodePanes.Release();
                CommandBars.Release();
                Windows.Release();
                AddIns.Release();
                base.Release(final);
            }
        }

        public override bool Equals(ISafeComWrapper<VB.VBE> other)
        {
            return IsEqualIfNull(other) || (other != null && other.Target.Version == Version);
        }

        public bool Equals(IVBE other)
        {
            return Equals(other as SafeComWrapper<VB.VBE>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }

        public bool IsInDesignMode
        {
            get { return VBProjects.All(project => project.Mode == EnvironmentMode.Design); }
        }

        public static void SetSelection(IVBProject vbProject, Selection selection, string name)
        {
            var components = vbProject.VBComponents;
            var component = components.SingleOrDefault(c => c.Name == name);
            if (component == null || component.IsWrappingNullReference)
            {
                return;
            }

            var module = component.CodeModule;
            var pane = module.CodePane;
            pane.Selection = selection;
        }

        private bool _hasfocus;
        private void WinEventProc(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, uint idObject, uint idChild, uint dwEventThread, uint dwmsEventTime)
        {
            if ((WinEvents.EventConstant)eventType == WinEvents.EventConstant.EVENT_SYSTEM_FOREGROUND)
            {
                _hasfocus = hwnd == _hwnd;
                if (hwnd == _hwnd)
                {
                    Debug.WriteLine("Focus {0} for hwnd {1:X8} ({2}).", _hasfocus ? "set" : "lost", hwnd.ToInt32(), hwnd.ToClassName());
                }
            }
            if (hwnd == _hwnd && (WinEvents.EventConstant) eventType == WinEvents.EventConstant.EVENT_OBJECT_DESTROY)
            {
                Debug.WriteLine("VBE out.");
                //TODO. Get rid of the access violation on the next line of the Output window.
            }
        }
    }
}
