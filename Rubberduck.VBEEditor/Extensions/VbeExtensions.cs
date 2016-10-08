using System.Linq;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Forms;
using Rubberduck.VBEditor.SafeComWrappers.VBA.Abstract;
using Rubberduck.VBEditor.VBEHost;

namespace Rubberduck.VBEditor.Extensions
{
    public static class VbeExtensions
    {
        /// <summary> Returns the type of Office Application that is hosting the VBE. </summary>
        public static IHostApplication HostApplication(this IVBE vbe)
        {
            var project = vbe.ActiveVBProject;
            {
                if (project.IsWrappingNullReference)
                {
                    const int ctlViewHost = 106;

                    var commandBars = vbe.CommandBars;
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
                                    result = new SolidWorksApp(vbe);
                                    break;
                                default:
                                    result = null;
                                    break;
                            }
                        }

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
            }

            return null;
        }

        /// <summary> Returns whether the host supports unit tests.</summary>
        public static bool HostSupportsUnitTests(this IVBE vbe)
        {
            var project = vbe.ActiveVBProject;
            {
                if (project.IsWrappingNullReference)
                {
                    const int ctlViewHost = 106;
                    var commandBars = vbe.CommandBars;
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
