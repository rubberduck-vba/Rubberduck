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
            var project = vbe.VBProjects.Cast<VBProject>()
                .SingleOrDefault(p => p.Protection != vbext_ProjectProtection.vbext_pp_locked
                                      && ReferenceEquals(p, vbProject));

            VBComponent component = null;
            if (project != null)
            {
                component = project.VBComponents.Cast<VBComponent>()
                    .SingleOrDefault(c => c.Name == name);
            }

            if (component == null)
            {
                return;
            }

            try
            {
                var codePane = wrapperFactory.Create(component.CodeModule.CodePane);
                codePane.Selection = selection;
            }
            catch (Exception e)
            {
            }
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

                CommandBarControl host_app_control = vbe.CommandBars.FindControl(MsoControlType.msoControlButton, ctl_view_host);

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
                        case "Microsoft Publisher":
                            return new PublisherApp();
                        case "AutoCAD":
                            return null; //TODO - Confirm the button caption
                        case "CorelDRAW":
                            return null;
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
                    case "Publisher":
                        return new PublisherApp();
					case "AutoCAD":
                        return new AutoCADApp();
                    case "CorelDRAW":
                        return new CorelDRAWApp(vbe);
                }
            }

            return null;
        }
    }
}