using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Rubberduck.VBEditor.Application;
using Rubberduck.VBEditor.SafeComWrappers.VB.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.VB.Enums;
using Rubberduck.VBEditor.SafeComWrappers.Office.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.Office.v12;
using Rubberduck.VBEditor.WindowsApi;
using VB6IA = Microsoft.VB6.Interop.VBIDE;

namespace Rubberduck.VBEditor.SafeComWrappers.VB.VB6
{
    public class VBE : SafeComWrapper<VB6IA.VBE>, IVBE
    {
        public VBE(VB6IA.VBE target)
            :base(target)
        {
        }

        public object HardReference => Target;

        public string Version => IsWrappingNullReference ? string.Empty : Target.Version;

        public ICodePane ActiveCodePane
        {
            get => new CodePane(IsWrappingNullReference ? null : Target.ActiveCodePane);
            set { if (!IsWrappingNullReference) Target.ActiveCodePane = (VB6IA.CodePane)value.Target; }
        }

        public IVBProject ActiveVBProject
        {
            get => new VBProject(IsWrappingNullReference ? null : Target.ActiveVBProject);
            set { if (!IsWrappingNullReference) Target.ActiveVBProject = (VB6IA.VBProject)value.Target; }
        }

        public IWindow ActiveWindow => new Window(IsWrappingNullReference ? null : Target.ActiveWindow);

        public IAddIns AddIns => new AddIns(IsWrappingNullReference ? null : Target.Addins);

        public ICodePanes CodePanes => new CodePanes(IsWrappingNullReference ? null : Target.CodePanes);

        public ICommandBars CommandBars => new CommandBars(IsWrappingNullReference ? null : Target.CommandBars);

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

        public IVBComponent SelectedVBComponent => new VBComponent(IsWrappingNullReference ? null : Target.SelectedVBComponent);

        public IVBProjects VBProjects => new VBProjects(IsWrappingNullReference ? null : Target.VBProjects);

        public IWindows Windows => new Windows(IsWrappingNullReference ? null : Target.Windows);

        //public override void Release(bool final = false)
        //{
        //    if (!IsWrappingNullReference)
        //    {
        //        VBProjects.Release();
        //        CodePanes.Release();
        //        CommandBars.Release();
        //        Windows.Release();
        //        AddIns.Release();
        //        base.Release(final);
        //    }
        //}

        public override bool Equals(ISafeComWrapper<VB6IA.VBE> other)
        {
            return IsEqualIfNull(other) || (other != null && other.Target.Version == Version);
        }

        public bool Equals(IVBE other)
        {
            return Equals(other as SafeComWrapper<VB6IA.VBE>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }

        public IHostApplication HostApplication()
        {
            return null;
        }

        public IWindow ActiveMDIChild()
        {
            const string mdiClientClass = "MDIClient";
            const int maxCaptionLength = 512;

            IntPtr mainWindow = (IntPtr)MainWindow.HWnd;

            IntPtr mdiClient = NativeMethods.FindWindowEx(mainWindow, IntPtr.Zero, mdiClientClass, string.Empty);

            IntPtr mdiChild = NativeMethods.GetTopWindow(mdiClient);
            StringBuilder mdiChildCaption = new StringBuilder();
            int captionLength = NativeMethods.GetWindowText(mdiChild, mdiChildCaption, maxCaptionLength);

            if (captionLength > 0)
            {
                try
                {
                    return Windows.FirstOrDefault(win => win.Caption == mdiChildCaption.ToString());
                }
                catch
                {
                }
            }
            return null;
        }

        public bool IsInDesignMode
        {
            get { return VBProjects.All(project => project.Mode == EnvironmentMode.Design); }
        }

        public void SetSelection(IVBProject vbProject, Selection selection, string name)
        {
            throw new NotImplementedException();
        }
    }
}
