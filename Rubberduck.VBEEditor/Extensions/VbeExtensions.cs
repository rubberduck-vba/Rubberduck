using System.Linq;
using Microsoft.Office.Core;
using Rubberduck.VBEditor.DisposableWrappers;
using Rubberduck.VBEditor.VBEHost;

namespace Rubberduck.VBEditor.Extensions
{
    public static class VbeExtensions
    {
        public static void SetSelection(this VBE vbe, VBProject vbProject, Selection selection, string name)
        {
            try
            {
                using (var component = vbProject.VBComponents.Cast<VBComponent>().SingleOrDefault(c => c.Name == name))
                {
                    if (component == null || component.IsWrappingNullReference)
                    {
                        return;
                    }

                    using (var module = component.CodeModule)
                    using (var pane = module.CodePane)
                    {
                        pane.SetSelection(selection);
                    }
                }
            }
            catch (WrapperMethodException)
            {
            }
        }

        public static bool IsInDesignMode(this VBE vbe)
        {
            return vbe.VBProjects.All(project => project.Mode == EnvironmentMode.Design);
        }

        public static CodeModuleSelection FindInstruction(this VBE vbe, QualifiedModuleName qualifiedModuleName, Selection selection)
        {
            var module = qualifiedModuleName.Component.CodeModule;
            if (module == null)
            {
                return null;
            }

            return new CodeModuleSelection(module, selection);
        }

        /// <summary> Returns the type of Office Application that is hosting the VBE. </summary>
        public static IHostApplication HostApplication(this VBE vbe)
        {
            if (vbe.ActiveVBProject == null)
            {
                const int ctlViewHost = 106;

                var hostAppControl = vbe.CommandBars.FindControl(MsoControlType.msoControlButton, ctlViewHost);

                if (hostAppControl == null)
                {
                    return null;
                }
                else
                {
                    switch (hostAppControl.Caption)
                    {
                        case "Microsoft Excel":
                            return new ExcelApp();
                        case "Microsoft Access":
                            return new AccessApp();
                        case "Microsoft Word":
                            return new WordApp();
                        case "Microsoft PowerPoint":
                            return new PowerPointApp();
                        case "Microsoft Outlook":
                            return new OutlookApp();
                        case "Microsoft Project":
                            return new ProjectApp();
                        case "Microsoft Publisher":
                            return new PublisherApp();
                        case "Microsoft Visio":
                            return new VisioApp();
                        case "AutoCAD":
                            return new AutoCADApp();
                        case "CorelDRAW":
                            return new CorelDRAWApp();
                        case "SolidWorks":
                            return new SolidWorksApp(vbe);
                    }
                }
                return null;
            }

            using (var project = vbe.ActiveVBProject)
            using (var references = project.References)
            {
                foreach (var reference in references.Where(reference => (reference.IsBuiltIn && reference.Name != "VBA") || (reference.Name == "AutoCAD")))
                {
                    switch (reference.Name)
                    {
                        case "Excel":
                            return new ExcelApp(vbe);
                        case "Access":
                            return new AccessApp();
                        case "Word":
                            return new WordApp(vbe);
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
                            return new CorelDRAWApp(vbe);
                        case "SolidWorks":
                            return new SolidWorksApp(vbe);
                    }
                }
            }

            return null;
        }

        /// <summary> Returns whether the host supports unit tests.</summary>
        public static bool HostSupportsUnitTests(this VBE vbe)
        {
            if (vbe.ActiveVBProject == null)
            {
                const int ctlViewHost = 106;

                var hostAppControl = vbe.CommandBars.FindControl(MsoControlType.msoControlButton, ctlViewHost);

                if (hostAppControl == null)
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

            using (var project = vbe.ActiveVBProject)
            {
                foreach (var reference in project.References
                    .Where(reference => (reference.IsBuiltIn && reference.Name != "VBA") || (reference.Name == "AutoCAD")))
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

            return false;
        }
    }
}
