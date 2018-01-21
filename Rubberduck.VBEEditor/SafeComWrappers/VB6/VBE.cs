using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Rubberduck.VBEditor.Application;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;
using Rubberduck.VBEditor.WindowsApi;
using VB = Microsoft.VB6.Interop.VBIDE;

namespace Rubberduck.VBEditor.SafeComWrappers.VB6
{
    public class VBE : SafeComWrapper<VB.VBE>, IVBE
    {
        public VBE(VB.VBE target, bool rewrapping = false)
            : base(target, rewrapping)
        {
        }

        public object HardReference => Target;

        public string Version => IsWrappingNullReference ? string.Empty : Target.Version;

        public ICodePane ActiveCodePane
        {
            get => new CodePane(IsWrappingNullReference ? null : Target.ActiveCodePane);
            set
            {
                if (!IsWrappingNullReference)
                {
                    Target.ActiveCodePane = (VB.CodePane) value.Target;
                }
            }
        }

        public IVBProject ActiveVBProject
        {
            get => new VBProject(IsWrappingNullReference ? null : Target.ActiveVBProject);
            set
            {
                if (!IsWrappingNullReference)
                {
                    Target.ActiveVBProject = (VB.VBProject) value.Target;
                }
            }
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

        public IHostApplication HostApplication()
        {
            return null;
        }

        public IWindow ActiveMDIChild()
        {
            const string mdiClientClass = "MDIClient";
            const int maxCaptionLength = 512;

            var mainWindow = (IntPtr) MainWindow.HWnd;

            var mdiClient = NativeMethods.FindWindowEx(mainWindow, IntPtr.Zero, mdiClientClass, string.Empty);

            var mdiChild = NativeMethods.GetTopWindow(mdiClient);
            var mdiChildCaption = new StringBuilder();
            var captionLength = NativeMethods.GetWindowText(mdiChild, mdiChildCaption, maxCaptionLength);

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
            get
            {
                var allInDesignMode = true;
                using (var projects = VBProjects)
                {
                    foreach (var project in projects)
                    {
                        allInDesignMode = allInDesignMode && project.Mode == EnvironmentMode.Design;
                        project.Dispose();
                        if (!allInDesignMode)
                        {
                            break;
                        }
                    }
                }
                return allInDesignMode;
            }
        }

        public int ProjectsCount
        {
            get
            {
                using (var projects = VBProjects)
                {
                    return projects.Count;
                }
            }
        }
    }
}
