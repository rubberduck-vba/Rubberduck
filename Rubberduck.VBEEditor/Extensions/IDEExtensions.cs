using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Rubberduck.VBEditor.Application;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.MSForms;
using Exception = System.Exception;

namespace Rubberduck.VBEditor.Extensions
{
    public static class VBEExtensions
    {
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

        /// <summary> Returns the type of Office Application that is hosting the VBE. </summary>
        public static IHostApplication HostApplication(this IVBE vbe)
        {
            var host = Path.GetFileName(System.Windows.Forms.Application.ExecutablePath).ToUpperInvariant();
            //This needs the VBE as a ctor argument.
            if (host.Equals("SLDWORKS.EXE"))
            {
                return new SolidWorksApp(vbe);
            }
            //The rest don't.
            if (HostAppMap.ContainsKey(host))
            {
                return (IHostApplication)Activator.CreateInstance(HostAppMap[host]);
            }

            //Guessing the above will work like 99.9999% of the time for supported applications.
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
            var host = Path.GetFileName(System.Windows.Forms.Application.ExecutablePath).ToUpperInvariant();
            if (HostAppMap.ContainsKey(host)) return true;
            //Guessing the above will work like 99.9999% of the time for supported applications.

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

    public static class VBProjectExtensions
    {
        /// <summary>
        /// Imports all source code files from target directory into project.
        /// </summary>
        /// <remarks>
        /// Only files with extensions "cls", "bas, "frm", and "doccls" are imported.
        /// It is the callers responsibility to remove any existing components prior to importing.
        /// </remarks>
        /// <param name="project"></param>
        /// <param name="filePath">Directory path containing the source files.</param>
        public static void ImportDocumentTypeSourceFiles(this IVBProject project, string filePath)
        {
            var dirInfo = new DirectoryInfo(filePath);

            var files = dirInfo.EnumerateFiles()
                                .Where(f => f.Extension == ComponentTypeExtensions.DocClassExtension);
            foreach (var file in files)
            {
                try
                {
                    project.VBComponents.ImportSourceFile(file.FullName);
                }
                catch (IndexOutOfRangeException) { }    // component didn't exist
            }
        }

        public static void LoadAllComponents(this IVBProject project, string filePath)
        {
            var dirInfo = new DirectoryInfo(filePath);

            var files = dirInfo.EnumerateFiles()
                                .Where(f => f.Extension == ComponentTypeExtensions.StandardExtension ||
                                            f.Extension == ComponentTypeExtensions.ClassExtension ||
                                            f.Extension == ComponentTypeExtensions.DocClassExtension ||
                                            f.Extension == ComponentTypeExtensions.FormExtension
                                            )
                                .ToList();

            var exceptions = new List<Exception>();

            foreach (var component in project.VBComponents)
            {
                try
                {
                    var name = component.Name;
                    project.VBComponents.RemoveSafely(component);

                    var file = files.SingleOrDefault(f => f.Name == name + f.Extension);
                    if (file != null)
                    {
                        try
                        {
                            project.VBComponents.ImportSourceFile(file.FullName);
                        }
                        catch (IndexOutOfRangeException)
                        {
                            exceptions.Add(new IndexOutOfRangeException(string.Format(VBEEditorText.NonexistentComponentErrorText, Path.GetFileNameWithoutExtension(file.FullName))));
                        }
                    }
                }
                catch (Exception ex)
                {
                    exceptions.Add(ex);
                }
            }

            foreach (var file in files)
            {
                try
                {
                    if (project.VBComponents.All(v => v.Name + file.Extension != file.Name))
                    {
                        try
                        {
                            project.VBComponents.ImportSourceFile(file.FullName);
                        }
                        catch (IndexOutOfRangeException)
                        {
                            exceptions.Add(new IndexOutOfRangeException(string.Format(VBEEditorText.NonexistentComponentErrorText, Path.GetFileNameWithoutExtension(file.FullName))));
                        }
                    }
                }
                catch (Exception ex)
                {
                    exceptions.Add(ex);
                }
            }

            if (exceptions.Count != 0)
            {
                throw new AggregateException(string.Empty, exceptions);
            }
        }
    }
}
