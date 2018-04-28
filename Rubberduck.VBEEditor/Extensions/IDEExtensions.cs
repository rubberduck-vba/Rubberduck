using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Exception = System.Exception;

namespace Rubberduck.VBEditor.Extensions
{


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
