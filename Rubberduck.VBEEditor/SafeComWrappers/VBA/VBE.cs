using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Rubberduck.VBEditor.Application;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.MSForms;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;
using Rubberduck.VBEditor.WindowsApi;
using VB = Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class VBE : SafeComWrapper<VB.VBE>, IVBE
    {
        public VBE(VB.VBE target)
            : base(target)
        {
        }

        public object HardReference
        {
            get { return Target; }
        }

        public string Version
        {
            get { return IsWrappingNullReference ? string.Empty : Target.Version; }
        }

        public ICodePane ActiveCodePane
        {
            get { return new CodePane(IsWrappingNullReference ? null : Target.ActiveCodePane); }
            set { if (!IsWrappingNullReference) Target.ActiveCodePane = (VB.CodePane)value.Target; }
        }

        public IVBProject ActiveVBProject
        {
            get { return new VBProject(IsWrappingNullReference ? null : Target.ActiveVBProject); }
            set { if (!IsWrappingNullReference) Target.ActiveVBProject = (VB.VBProject)value.Target; }
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

        //public override void Release(bool final = false)
        //{
        //    if (!IsWrappingNullReference)
        //    {
        //        VBProjects.Release();
        //        CodePanes.Release();
        //        //CommandBars.Release();
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
            //This needs the VBE as a ctor argument.
            if (host.Equals("SLDWORKS.EXE"))
            {
                return new SolidWorksApp(this);
            }
            //The rest don't.
            if (HostAppMap.ContainsKey(host))
            {
                return (IHostApplication)Activator.CreateInstance(HostAppMap[host]);
            }

            //Guessing the above will work like 99.9999% of the time for supported applications.
            var project = ActiveVBProject;
            {
                if (project.IsWrappingNullReference)
                {
                    const int ctlViewHost = 106;

                    var commandBars = CommandBars;
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
                                    result = new ExcelApp();
                                    break;
                                case "Microsoft Access":
                                    result = new AccessApp();
                                    break;
                                case "Microsoft Word":
                                    result = new WordApp();
                                    break;
                                case "Microsoft PowerPoint":
                                    result = new PowerPointApp();
                                    break;
                                case "Microsoft Outlook":
                                    result = new OutlookApp();
                                    break;
                                case "Microsoft Project":
                                    result = new ProjectApp();
                                    break;
                                case "Microsoft Publisher":
                                    result = new PublisherApp();
                                    break;
                                case "Microsoft Visio":
                                    result = new VisioApp();
                                    break;
                                case "AutoCAD":
                                    result = new AutoCADApp();
                                    break;
                                case "CorelDRAW":
                                    result = new CorelDRAWApp();
                                    break;
                                case "SolidWorks":
                                    result = new SolidWorksApp(this);
                                    break;
                                default:
                                    result = null;
                                    break;
                            }
                        }

                        _host = result;
                        return result;
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
                                return new AccessApp();
                            case "Word":
                                return new WordApp(this);
                            case "PowerPoint":
                                return new PowerPointApp();
                            case "Outlook":
                                return new OutlookApp();
                            case "MSProject":
                                return new ProjectApp();
                            case "Publisher":
                                return new PublisherApp();
                            case "Visio":
                                return new VisioApp();
                            case "AutoCAD":
                                return new AutoCADApp();
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

        /// <summary> Returns whether the host supports unit tests.</summary>
        public bool HostSupportsUnitTests()
        {
            var host = Path.GetFileName(System.Windows.Forms.Application.ExecutablePath).ToUpperInvariant();
            if (HostAppMap.ContainsKey(host)) return true;
            //Guessing the above will work like 99.9999% of the time for supported applications.

            var project = ActiveVBProject;
            {
                if (project.IsWrappingNullReference)
                {
                    const int ctlViewHost = 106;
                    var commandBars = CommandBars;
                    var hostAppControl = commandBars.FindControl(ControlType.Button, ctlViewHost);
                    {
                        if (hostAppControl.IsWrappingNullReference)
                        {
                            return false;
                        }

                        switch (hostAppControl.Caption)
                        {
                            case "Microsoft Excel":
                            case "Microsoft Access":
                            case "Microsoft Word":
                            case "Microsoft PowerPoint":
                            case "Microsoft Outlook":
                            case "Microsoft Project":
                            case "Microsoft Publisher":
                            case "Microsoft Visio":
                            case "AutoCAD":
                            case "CorelDRAW":
                            case "SolidWorks":
                                return true;
                            default:
                                return false;
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
                            case "Access":
                            case "Word":
                            case "PowerPoint":
                            case "Outlook":
                            case "MSProject":
                            case "Publisher":
                            case "Visio":
                            case "AutoCAD":
                            case "CorelDRAW":
                            case "SolidWorks":
                                return true;
                        }
                    }
                }
            }
            return false;
        }
    }
}
