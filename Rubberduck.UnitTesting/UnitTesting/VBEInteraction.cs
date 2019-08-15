
using Microsoft.Win32;
using NLog;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources.Registration;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace Rubberduck.UnitTesting
{ 
    // FIXME litter some logging around here
    internal class VBEInteraction : IVBEInteraction
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private readonly Version _rubberduckVersion;
        private readonly IVBETypeLibsAPI _typeLibsApi;

        public VBEInteraction(IVBETypeLibsAPI typeLibsApi, Version rubberduckVersion)
        {
            _typeLibsApi = typeLibsApi;
            _rubberduckVersion = rubberduckVersion;
        }

        public void RunDeclarations(ITypeLibWrapper typeLib, IEnumerable<Declaration> declarations)
        {
            foreach (var declaration in declarations)
            {
                _typeLibsApi.ExecuteCode(typeLib, declaration.QualifiedModuleName.ComponentName,
                    declaration.QualifiedName.MemberName);
            }
        }

        public void RunTestMethod(ITypeLibWrapper typeLib, TestMethod test, EventHandler<AssertCompletedEventArgs> assertCompletionHandler, out long duration)
        {
            AssertHandler.OnAssertCompleted += assertCompletionHandler;
            var stopwatch = new Stopwatch();
            try
            {
                var testDeclaration = test.Declaration;

                stopwatch.Start();
                _typeLibsApi.ExecuteCode(typeLib, testDeclaration.ComponentName, testDeclaration.QualifiedName.MemberName);
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
            

        public void EnsureProjectReferencesUnitTesting(IVBProject project)
        {
            if (project == null || project.IsWrappingNullReference) { return; }
            var libFolder = IntPtr.Size == 8 ? "win64" : "win32";
            const string libGuid = RubberduckGuid.RubberduckTypeLibGuid;
            var pathKey = Registry.ClassesRoot.OpenSubKey(
                $@"TypeLib\{{{libGuid}}}\{_rubberduckVersion.Major}.{_rubberduckVersion.Minor}\0\{libFolder}");

            if (pathKey != null)
            {
                var referencePath = pathKey.GetValue(string.Empty, string.Empty) as string;
                string name = null;

                if (!string.IsNullOrWhiteSpace(referencePath))
                {
                    var tlbKey =
                        Registry.ClassesRoot.OpenSubKey(
                            $@"TypeLib\{{{libGuid}}}\{_rubberduckVersion.Major}.{_rubberduckVersion.Minor}");

                    if(tlbKey != null)
                    {
                        name = tlbKey.GetValue(string.Empty, string.Empty) as string;
                        tlbKey.Dispose();
                    }
                }

                if (string.IsNullOrWhiteSpace(referencePath) || string.IsNullOrWhiteSpace(name))
                {
                    throw new InvalidOperationException(
                        "Cannot locate the tlb in the registry or the entry may be corrupted. Therefore early binding is not possible");
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
                        // AddFromFile returns a new wrapped reference so we must 
                        // ensure it is disposed properly.
                        using (references.AddFromFile(referencePath)) { }
                    }
                }

                pathKey.Dispose();
            }
        }

        private static IReference FindReferenceByName(IReferences references, string name)
        {
            foreach (var reference in references)
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
