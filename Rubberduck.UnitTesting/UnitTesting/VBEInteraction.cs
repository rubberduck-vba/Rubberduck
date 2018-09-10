
using Microsoft.Win32;
using NLog;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources.Registration;
using Rubberduck.VBEditor.ComManagement.TypeLibs;
using Rubberduck.VBEditor.ComManagement.TypeLibsAPI;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;

namespace Rubberduck.UnitTesting
{ 
    // FIXME litter some logging around here
    public class VBEInteraction
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        internal static void RunDeclarations(IVBETypeLibsAPI typeLibApi, ITypeLibWrapper typeLib, IEnumerable<Declaration> declarations)
        {
            foreach (var declaration in declarations)
            {
                typeLibApi.ExecuteCode(typeLib, declaration.QualifiedModuleName.ComponentName,
                    declaration.QualifiedName.MemberName);
            }
        }

        internal static void RunTestMethod(IVBETypeLibsAPI tlApi, ITypeLibWrapper typeLib, TestMethod test, EventHandler<AssertCompletedEventArgs> assertCompletionHandler, out long duration)
        {
            AssertHandler.OnAssertCompleted += assertCompletionHandler;
            var stopwatch = new Stopwatch();
            try
            {
                var testDeclaration = test.Declaration;

                stopwatch.Start();
                tlApi.ExecuteCode(typeLib, testDeclaration.ComponentName, testDeclaration.QualifiedName.MemberName);
                stopwatch.Stop();

                duration = stopwatch.ElapsedMilliseconds;
            }
            catch (Exception)
            {
                stopwatch.Stop();
                duration = stopwatch.ElapsedMilliseconds;
                throw;
            }
            finally
            {
                AssertHandler.OnAssertCompleted -= assertCompletionHandler;
            }
        }
            

        public static void EnsureProjectReferencesUnitTesting(IVBProject project)
        {
            if (project == null || project.IsWrappingNullReference) { return; }
            var libFolder = IntPtr.Size == 8 ? "win64" : "win32";
            // FIXME: This assumes the current assembly is same major/minor as the TLB!!!
            var libVersion = Assembly.GetExecutingAssembly().GetName().Version;
            const string libGuid = RubberduckGuid.RubberduckTypeLibGuid;
            var pathKey = Registry.ClassesRoot.OpenSubKey($@"TypeLib\{{{libGuid}}}\{libVersion.Major}.{libVersion.Minor}\0\{libFolder}");

            var referencePath = pathKey?.GetValue(string.Empty, string.Empty) as string;
            string name = null;

            if (!string.IsNullOrWhiteSpace(referencePath))
            {
                var tlbKey =
                    Registry.ClassesRoot.OpenSubKey($@"TypeLib\{{{libGuid}}}\{libVersion.Major}.{libVersion.Minor}");

                name = tlbKey?.GetValue(string.Empty, string.Empty) as string;
            }

            if (string.IsNullOrWhiteSpace(referencePath) || string.IsNullOrWhiteSpace(name))
            {
                throw new InvalidOperationException("Cannot locate the tlb in the registry or the entry may be corrupted. Therefore early binding is not possible");
            }

            using (var references = project.References)
            {
                var reference = FindReferenceByName(references, name);
                if (reference != null)
                {
                    references.Remove(reference);
                    reference.Dispose();
                }

                if (!ReferenceWithPathExists(references, referencePath))
                {
                    references.AddFromFile(referencePath);
                }
            }
        }

        private static IReference FindReferenceByName(IReferences refernences, string name)
        {
            foreach (var reference in refernences)
            {
                if (reference.Name == name)
                {
                    return reference;
                }
                reference.Dispose();
            }

            return null;
        }

        private static bool ReferenceWithPathExists(IReferences refereneces, string path)
        {
            foreach (var reference in refereneces)
            {
                var referencePath = reference.FullPath;
                reference.Dispose();
                if (referencePath == path)
                {
                    return true;
                }
            }
            return false;
        }

    }
}
