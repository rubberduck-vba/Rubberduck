using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.Extensions
{
    public static class ProjectExtensions
    {
        public static string ProjectName(this VBProject project)
        {
            var projectName = string.Empty;
            try
            {
                projectName = project.Name;
                if (project.Protection == vbext_ProjectProtection.vbext_pp_none)
                {
                    var documentModule = project
                        .VBComponents.Cast<VBComponent>()
                        .FirstOrDefault(item => item.Type == vbext_ComponentType.vbext_ct_Document);

                    if (documentModule != null)
                    {
                        var hostDocumentNameProperty = documentModule.Properties.Item("Name");
                        if (hostDocumentNameProperty != null)
                        {
                            projectName += string.Format(" ({0})", hostDocumentNameProperty.Value);
                        }
                    }
                }
            }
            catch (COMException exception)
            {
                Debug.WriteLine(exception);
            }
            return projectName;
        }

        public static string ProjectName(this VBComponent component)
        {
            var project = component.Collection.Parent;
            return project.ProjectName();
        }

        public static IEnumerable<string> ComponentNames(this VBProject project)
        {
            foreach (VBComponent component in project.VBComponents)
            {
                yield return component.Name;
            }
        }

        /// <summary>
        /// Exports all code modules in the VbProject to a destination directory. Files are given the same name as their parent code Module name and file extensions are based on what type of code Module it is.
        /// </summary>
        /// <param name="project">The <see cref="VBProject"/> to be exported to source files.</param>
        /// <param name="directoryPath">The destination directory path.</param>
        public static void ExportSourceFiles(this VBProject project, string directoryPath)
        {
            foreach (VBComponent component in project.VBComponents)
            {
                component.ExportAsSourceFile(directoryPath);
            }
        }

        /// <summary>
        /// Removes All VbComponents from the VbProject.
        /// </summary>
        /// <remarks>
        /// Document type Components cannot be physically removed from a project through the VBE.
        /// Instead, the code will simply be deleted from the code Module.
        /// </remarks>
        /// <param name="project"></param>
        public static void RemoveAllComponents(this VBProject project)
        {
            foreach (VBComponent component in project.VBComponents)
            {
                project.VBComponents.RemoveSafely(component);
            }
        }

        /// <summary>
        /// Imports all source code files from target directory into project.
        /// </summary>
        /// <remarks>
        /// Only files with extensions "cls", "bas, "frm", and "doccls" are imported.
        /// It is the callers responsibility to remove any existing components prior to importing.
        /// </remarks>
        /// <param name="project"></param>
        /// <param name="filePath">Directory path containing the source files.</param>
        public static void ImportSourceFiles(this VBProject project, string filePath)
        {
            var dirInfo = new DirectoryInfo(filePath);

            var files = dirInfo.EnumerateFiles()
                                .Where(f => f.Extension == VBComponentExtensions.StandardExtension ||
                                            f.Extension == VBComponentExtensions.ClassExtension ||
                                            f.Extension == VBComponentExtensions.DocClassExtension ||
                                            f.Extension == VBComponentExtensions.FormExtension
                                            );
            foreach (var file in files)
            {
                project.VBComponents.ImportSourceFile(file.FullName);
            }
        }
    }
}
