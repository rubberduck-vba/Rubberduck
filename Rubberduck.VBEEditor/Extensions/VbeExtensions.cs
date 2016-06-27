using System;
using System.Linq;
using Microsoft.Vbe.Interop;
using Microsoft.Office.Core;
using Rubberduck.VBEditor.VBEHost;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.VBEditor.Extensions
{
    public static class VbeExtensions
    {
        public static void SetSelection(this VBE vbe, VBProject vbProject, Selection selection, string name,
            ICodePaneWrapperFactory wrapperFactory)
        {
            try
            {
                var component = vbProject.VBComponents.Cast<VBComponent>().SingleOrDefault(c => c.Name == name);
                if (component == null)
                {
                    return;
                }

                var codePane = wrapperFactory.Create(component.CodeModule.CodePane);
                codePane.Selection = selection;
            }
            catch (Exception e)
            {
            }
        }

        public static bool IsInDesignMode(this VBE vbe)
        {
            return vbe.VBProjects.Cast<VBProject>().All(project => project.Mode == vbext_VBAMode.vbext_vm_Design);
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
                const int ctl_view_host = 106;

                var host_app_control = vbe.CommandBars.FindControl(MsoControlType.msoControlButton, ctl_view_host);

                if (host_app_control == null)
                {
                    return null;
                }
                else
                {
                    switch (host_app_control.Caption)
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
                    }
                }
                return null;
            }

            foreach (var reference in vbe.ActiveVBProject.References.Cast<Reference>()
                .Where(reference => (reference.BuiltIn && reference.Name != "VBA") || (reference.Name == "AutoCAD")))
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
                        return true;
                    default:
                        return false;
                }
            }

            foreach (var reference in vbe.ActiveVBProject.References.Cast<Reference>()
                .Where(reference => (reference.BuiltIn && reference.Name != "VBA") || (reference.Name == "AutoCAD")))
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
                        return true;
                }
            }

            return false;
        }
    }
}
