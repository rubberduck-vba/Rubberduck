using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UnitTesting
{
    public static class UnitTestUtils // todo: reimplement using state.DeclarationFinder
    {
        public static IEnumerable<TestMethod> GetAllTests(IVBE vbe, RubberduckParserState state)
        {
            return GetTestModuleProcedures(state)
                    .Where(item => IsTestMethod(state, item))
                    .Select(item => new TestMethod(item, vbe));
        }

        public static IEnumerable<TestMethod> GetTests(this IVBComponent component, IVBE vbe, RubberduckParserState state)
        {
            if (component == null || component.IsWrappingNullReference)
            {
                return Enumerable.Empty<TestMethod>();
            }

            // apparently, sometimes it thinks the components are different but knows the modules are the same
            // if the modules are the same, then the component is the same as far as we are concerned
            return GetAllTests(vbe, state)
                    .Where(test => test.Declaration.QualifiedName.QualifiedModuleName.Component.CodeModule.Equals(component.CodeModule));
        }

        public static bool IsTestMethod(RubberduckParserState state, Declaration item)
        {
            return !state.AllUserDeclarations.Any(
                    d => d.DeclarationType == DeclarationType.Parameter && Equals(d.ParentScopeDeclaration, item)) &&
                item.Accessibility == Accessibility.Public &&
                item.Annotations.Any(a => a.AnnotationType == AnnotationType.TestMethod);
        }

        public static IEnumerable<Declaration> FindModuleInitializeMethods(this QualifiedModuleName module, RubberduckParserState state)
        {
            return GetTestModuleProcedures(state)
                    .Where(m =>
                            m.QualifiedName.QualifiedModuleName == module &&
                            m.Annotations.Any(a => a.AnnotationType == AnnotationType.ModuleInitialize));
        }

        public static IEnumerable<Declaration> FindModuleCleanupMethods(this QualifiedModuleName module, RubberduckParserState state)
        {
            return GetTestModuleProcedures(state)
                    .Where(m =>
                            m.QualifiedName.QualifiedModuleName == module &&
                            m.Annotations.Any(a => a.AnnotationType == AnnotationType.ModuleCleanup));
        }

        public static IEnumerable<Declaration> FindTestInitializeMethods(this QualifiedModuleName module, RubberduckParserState state)
        {
            return GetTestModuleProcedures(state)
                    .Where(m =>
                            m.QualifiedName.QualifiedModuleName == module &&
                            m.Annotations.Any(a => a.AnnotationType == AnnotationType.TestInitialize));
        }

        public static IEnumerable<Declaration> FindTestCleanupMethods(this QualifiedModuleName module, RubberduckParserState state)
        {
            return GetTestModuleProcedures(state)
                    .Where(m =>
                            m.QualifiedName.QualifiedModuleName == module &&
                            m.Annotations.Any(a => a.AnnotationType == AnnotationType.TestCleanup));
        }

        private static IEnumerable<Declaration> GetTestModuleProcedures(RubberduckParserState state)
        {
            var procedures = state.AllUserDeclarations.Where(item => item.DeclarationType == DeclarationType.Procedure);

            return procedures.Where(item =>
                        item.ParentDeclaration.DeclarationType == DeclarationType.ProceduralModule &&
                        item.ParentDeclaration.Annotations.Any(a => a.AnnotationType == AnnotationType.TestModule));
        }

        public static IEnumerable<Declaration> GetTestModules(this RubberduckParserState state)
        {
            return state.AllUserDeclarations.Where(item =>
                        item.DeclarationType == DeclarationType.ProceduralModule &&
                        item.Annotations.Any(a => a.AnnotationType == AnnotationType.TestModule));
        }
    }
}