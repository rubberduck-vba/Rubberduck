using System;
using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UnitTesting
{
    public interface IVBEInteraction
    {
        void EnsureProjectReferencesUnitTesting(IVBProject project);
        void RunDeclarations(ITypeLibWrapper typeLib, IEnumerable<Declaration> declarations);
        void RunTestMethod(ITypeLibWrapper typeLib, TestMethod test, EventHandler<AssertCompletedEventArgs> assertCompletionHandler, out long duration);
    }
}