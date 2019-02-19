using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.Office12;
using Rubberduck.VBEditor.SourceCodeHandling;
using Rubberduck.VBEditor.VBA;
using Rubberduck.VBEditor.WindowsApi;
using VB = Microsoft.Vbe.Interop;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class VBE : SafeComWrapper<VB.VBE>, IVBE
    {
        public VBE(VB.VBE target, bool rewrapping = false)
            : base(target, rewrapping)
        {
            TempSourceFileHandler = new TempSourceFileHandler();
        }

        public VBEKind Kind => VBEKind.Hosted;
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


        private static readonly Dictionary<string, Type> HostAppMap = new Dictionary<string, Type>
        {
            {"EXCEL.EXE", typeof(ExcelApp)},
            {"WINWORD.EXE", typeof(WordApp)},
            {"MSACCESS.EXE", typeof(AccessApp)},
            {"POWERPNT.EXE", typeof(PowerPointApp)},
            {"OUTLOOK.EXE", typeof(OutlookApp)},
            {"WINPROJ.EXE", typeof(ProjectApp)},
            {"MSPUB.EXE", typeof(PublisherApp)},
            {"VISIO.EXE", typeof(VisioApp)},
            {"ACAD.EXE", typeof(AutoCADApp)},
            {"CORELDRW.EXE", typeof(CorelDRAWApp)},
            {"SLDWORKS.EXE", typeof(SolidWorksApp)},
        };

        private static IHostApplication _host;

        /// <summary> Returns the type of Office Application that is hosting the VBE. </summary>
        public IHostApplication HostApplication()
        {
            // host app isn't going to change between calls. cache it.
            if (_host != null)
            {
                return _host;
            }

            var host = Path.GetFileName(System.Windows.Forms.Application.ExecutablePath).ToUpperInvariant();

            if (HostAppMap.ContainsKey(host))
            {
                return (IHostApplication)Activator.CreateInstance(HostAppMap[host], this);
            }

            //Guessing the above will work like 99.9999% of the time for supported applications.
            using (var project = ActiveVBProject)
            {
                if (project.IsWrappingNullReference)
                {
                    const int ctlViewHost = 106;

                    using (var commandBars = CommandBars)
                    {
                        var hostAppControl = commandBars.FindControl(ControlType.Button, ctlViewHost);
                        {
                            IHostApplication result;
                            if (hostAppControl.IsWrappingNullReference)
                            {
                                result = null;
                            }
                            else
                            {
                                switch (hostAppControl.Caption)
                                {
                                    case "Microsoft Excel":
                                        result = new ExcelApp(this);
                                        break;
                                    case "Microsoft Access":
                                        result = new AccessApp(this);
                                        break;
                                    case "Microsoft Word":
                                        result = new WordApp(this);
                                        break;
                                    case "Microsoft PowerPoint":
                                        result = new PowerPointApp(this);
                                        break;
                                    case "Microsoft Outlook":
                                        result = new OutlookApp(this);
                                        break;
                                    case "Microsoft Project":
                                        result = new ProjectApp(this);
                                        break;
                                    case "Microsoft Publisher":
                                        result = new PublisherApp(this);
                                        break;
                                    case "Microsoft Visio":
                                        result = new VisioApp(this);
                                        break;
                                    case "AutoCAD":
                                        result = new AutoCADApp(this);
                                        break;
                                    case "CorelDRAW":
                                        result = new CorelDRAWApp(this);
                                        break;
                                    case "SolidWorks":
                                        result = new SolidWorksApp(this);
                                        break;
                                    default:
                                        result = null;
                                        break;
                                }
                            }
                        }
                    }
                }

                var references = project.References;
                {
                    foreach (var reference in references.Where(reference => (reference.IsBuiltIn && reference.Name != "VBA") || (reference.Name == "AutoCAD")))
                    {
                        switch (reference.Name)
                        {
                            case "Excel":
                                return new ExcelApp(this);
                            case "Access":
                                return new AccessApp(this);
                            case "Word":
                                return new WordApp(this);
                            case "PowerPoint":
                                return new PowerPointApp(this);
                            case "Outlook":
                                return new OutlookApp(this);
                            case "MSProject":
                                return new ProjectApp(this);
                            case "Publisher":
                                return new PublisherApp(this);
                            case "Visio":
                                return new VisioApp(this);
                            case "AutoCAD":
                                return new AutoCADApp(this);
                            case "CorelDRAW":
                                return new CorelDRAWApp(this);
                            case "SolidWorks":
                                return new SolidWorksApp(this);
                        }
                    }
                }
            }

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

