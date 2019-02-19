using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.Office8;
using Rubberduck.VBEditor.SourceCodeHandling;
using Rubberduck.VBEditor.VB6;
using Rubberduck.VBEditor.WindowsApi;
using VB = Microsoft.Vbe.Interop.VB6;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VB6
{
    public class VBE : SafeComWrapper<VB.VBE>, IVBE
    {
        public VBE(VB.VBE target, bool rewrapping = false)
            : base(target, rewrapping)
        {
            TempSourceFileHandler = new ExternalFileTempSourceFileHandlerEmulator();
        }

        public VBEKind Kind => VBEKind.Standalone;
        public object HardReference => Target;
        public ITempSourceFileHandler TempSourceFileHandler { get; }

        public string Version => IsWrappingNullReference ? string.Empty : Target.Version;

        public ICodePane ActiveCodePane
        {
            get => new CodePane(IsWrappingNullReference ? null : Target.ActiveCodePane);
            set => SetActiveCodePane(value);
        }

        private void SetActiveCodePane(ICodePane codePane)
        {
            if (IsWrappingNullReference || !(codePane is CodePane pane))
            {
                return;
            }

            Target.ActiveCodePane = pane.Target;
            ForceFocus(codePane);
        }

        private void ForceFocus(ICodePane codePane)
        {
            if (codePane.IsWrappingNullReference)
            {
                return;
            }

            codePane.Show();

            using (var mainWindow = MainWindow)
            using (var window = codePane.Window)
            {
                var mainWindowHandle = mainWindow.Handle();
                var handle = mainWindow.Handle().FindChildWindow(window.Caption);

                if (handle != IntPtr.Zero)
                {
                    NativeMethods.ActivateWindow(handle, mainWindowHandle);
                }
                else
                {
                    _logger.Debug("VBE.ForceFocus() failed to get a handle on the MainWindow.");
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
                    Target.ActiveVBProject = (VB.VBProject)value.Target;
                }
            }
        }

        public IWindow ActiveWindow => new Window(IsWrappingNullReference ? null : Target.ActiveWindow);

        public IAddIns AddIns => new AddIns(IsWrappingNullReference ? null : Target.Addins);

        public ICodePanes CodePanes => new CodePanes(IsWrappingNullReference ? null : Target.CodePanes);

        public ICommandBars CommandBars => new CommandBars(IsWrappingNullReference ? null : Target.CommandBars, this);

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

        public IEvents Events => new Events(IsWrappingNullReference ? null : Target.Events);
        
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

        public static void SetSelection(IVBProject vbProject, Selection selection, string name)
        {
            using (var components = vbProject.VBComponents)
            {
                using (var component = components.SingleOrDefault(c => ComponentHasName(c, name)))
                {
                    if (component == null || component.IsWrappingNullReference)
                    {
                        return;
                    }

                    using (var module = component.CodeModule)
                    {
                        using (var pane = module.CodePane)
                        {
                            pane.Selection = selection;
                        }
                    }
                }
            }
        }

        private static bool ComponentHasName(IVBComponent c, string name)
        {
            var sameName = c.Name == name;
            if (!sameName)
            {
                c.Dispose();
            }
            return sameName;
        }


        public IHostApplication HostApplication()
        {
            return null;
        }

        /// <summary> Returns the topmost MDI child window. </summary>
        public IWindow ActiveMDIChild()
        {
            const string mdiClientClass = "MDIClient";
            const int maxCaptionLength = 512;

            var mainWindow = (IntPtr)MainWindow.HWnd;

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
        
        public QualifiedSelection? GetActiveSelection()
        {
            using (var activePane = ActiveCodePane)
            {
                if (activePane == null || activePane.IsWrappingNullReference)
                {
                    return null;
                }

                return activePane.GetQualifiedSelection();
            }
        }

        protected override void Dispose(bool disposing) => base.Dispose(disposing);
    }
}

