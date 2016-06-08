using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.UnitTesting
{
    public static class UnitTestHelpers
    {
        public static IEnumerable<TestMethod> GetAllTests(VBE vbe, RubberduckParserState state)
        {
            return GetTestModuleProcedures(state)
                    .Where(item => IsTestMethod(state, item))
                    .Select(item => new TestMethod(item, vbe));
        }

        public static IEnumerable<TestMethod> GetTests(this VBComponent component, VBE vbe, RubberduckParserState state)
        {
            // apparently, sometimes it thinks the components are different but knows the modules are the same
            // if the modules are the same, then the component is the same as far as we are concerned
            return GetAllTests(vbe, state)
                    .Where(test => test.QualifiedMemberName.QualifiedModuleName.Component.CodeModule == component.CodeModule);
        }

        public static bool IsTestMethod(RubberduckParserState state, Declaration item)
        {
            return !state.AllUserDeclarations.Any(
                    d => d.DeclarationType == DeclarationType.Parameter && d.ParentScopeDeclaration == item) &&
                item.Accessibility == Accessibility.Public &&
                item.Annotations.Any(a => a.AnnotationType == AnnotationType.TestMethod);
        }

        public static IEnumerable<QualifiedMemberName> FindModuleInitializeMethods(this QualifiedModuleName module, RubberduckParserState state)
        {
            return GetTestModuleProcedures(state)
                    .Where(m =>
                            m.QualifiedName.QualifiedModuleName == module &&
                            m.Annotations.Any(a => a.AnnotationType == AnnotationType.ModuleInitialize))
                    .Select(s => s.QualifiedName);
        }

        public static IEnumerable<QualifiedMemberName> FindModuleCleanupMethods(this QualifiedModuleName module, RubberduckParserState state)
        {
            return GetTestModuleProcedures(state)
                    .Where(m =>
                            m.QualifiedName.QualifiedModuleName == module &&
                            m.Annotations.Any(a => a.AnnotationType == AnnotationType.ModuleCleanup))
                    .Select(s => s.QualifiedName);
        }

        public static IEnumerable<QualifiedMemberName> FindTestInitializeMethods(this QualifiedModuleName module, RubberduckParserState state)
        {
            return GetTestModuleProcedures(state)
                    .Where(m =>
                            m.QualifiedName.QualifiedModuleName == module &&
                            m.Annotations.Any(a => a.AnnotationType == AnnotationType.TestInitialize))
                    .Select(s => s.QualifiedName);
        }

        public static IEnumerable<QualifiedMemberName> FindTestCleanupMethods(this QualifiedModuleName module, RubberduckParserState state)
        {
            return GetTestModuleProcedures(state)
                    .Where(m =>
                            m.QualifiedName.QualifiedModuleName == module &&
                            m.Annotations.Any(a => a.AnnotationType == AnnotationType.TestCleanup))
                    .Select(s => s.QualifiedName);
        }

        private static IEnumerable<Declaration> GetTestModuleProcedures(RubberduckParserState state)
        {
            var procedures = state.AllUserDeclarations.Where(item => item.DeclarationType == DeclarationType.Procedure);

            return procedures.Where(item =>
                        item.ParentDeclaration.DeclarationType == DeclarationType.ProceduralModule &&
                        item.ParentDeclaration.Annotations.Any(a => a.AnnotationType == AnnotationType.TestModule));
        }
    }
}