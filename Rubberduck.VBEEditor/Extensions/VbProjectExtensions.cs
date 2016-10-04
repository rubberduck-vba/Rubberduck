using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Rubberduck.VBEditor.DisposableWrappers;
using Rubberduck.VBEditor.DisposableWrappers.VBA;

namespace Rubberduck.VBEditor.Extensions
{
    public static class ProjectExtensions
    {
        public static string AssignProjectId(this VBProject project)
        {
            //assign a hashcode if no helpfile is present
            if (string.IsNullOrEmpty(project.HelpFile))
            {
                project.HelpFile = project.GetHashCode().ToString();
            }

            //loop until the helpfile is unique for this host session
            while (!IsProjectIdUnique(project.HelpFile, project.VBE))
            {
                project.HelpFile = (project.GetHashCode() ^ project.HelpFile.GetHashCode()).ToString();
            }

            return project.HelpFile;
        }

        private static bool IsProjectIdUnique(string id, VBE vbe)
        {
            var projectsWithId = 0;

            foreach (var project in vbe.VBProjects.Cast<VBProject>())
            {
                if (project.HelpFile == id)
                {
                    projectsWithId++;
                }
            }

            return projectsWithId == 1;
        }

        public static IEnumerable<VBProject> UnprotectedProjects(this VBProjects projects)
        {
            return projects.Cast<VBProject>().Where(project => project.Protection == ProjectProtection.Unprotected);
        }

        public static IEnumerable<string> ComponentNames(this VBProject project)
        {
            return project.VBComponents.Cast<VBComponent>().Select(component => component.Name);
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


        /// <summary>
        /// Imports all source code files from target directory into project.
        /// </summary>
        /// <remarks>
        /// Only files with extensions "cls", "bas, "frm", and "doccls" are imported.
        /// It is the callers responsibility to remove any existing components prior to importing.
        /// </remarks>
        /// <param name="project"></param>
        /// <param name="filePath">Directory path containing the source files.</param>
        public static void ImportDocumentTypeSourceFiles(this VBProject project, string filePath)
        {
            var dirInfo = new DirectoryInfo(filePath);

            var files = dirInfo.EnumerateFiles()
                                .Where(f => f.Extension == VBComponentExtensions.DocClassExtension);
            foreach (var file in files)
            {
                try
                {
                    project.VBComponents.ImportSourceFile(file.FullName);
                }
                catch (IndexOutOfRangeException) { }    // component didn't exist
            }
        }

        public static void LoadAllComponents(this VBProject project, string filePath)
        {
            var dirInfo = new DirectoryInfo(filePath);

            var files = dirInfo.EnumerateFiles()
                                .Where(f => f.Extension == VBComponentExtensions.StandardExtension ||
                                            f.Extension == VBComponentExtensions.ClassExtension ||
                                            f.Extension == VBComponentExtensions.DocClassExtension ||
                                            f.Extension == VBComponentExtensions.FormExtension
                                            );

            var exceptions = new List<Exception>();

            foreach (VBComponent component in project.VBComponents)
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
                    if (project.VBComponents.OfType<VBComponent>().All(v => v.Name + file.Extension != file.Name))
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
