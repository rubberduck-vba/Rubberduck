using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Annotations.Concrete;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UnitTesting
{
    // FIXME make internal. Nobody outside of RD.UnitTesting needs this! 
    public static class TestDiscovery // todo: reimplement using state.DeclarationFinder 
    {
        public static IEnumerable<TestMethod> GetAllTests(RubberduckParserState state)
        {
            return GetTestModuleProcedures(state)
                    .Where(item => IsTestMethod(state, item))
                    .Select(item => new TestMethod(item));
        }

        public static IEnumerable<TestMethod> GetTests(IVBE vbe, IVBComponent component, RubberduckParserState state)
        {
            if (component == null || component.IsWrappingNullReference)
            {
                return Enumerable.Empty<TestMethod>();
            }

            // apparently, sometimes it thinks the components are different but knows the modules are the same
            // if the modules are the same, then the component is the same as far as we are concerned
            return GetAllTests(state)
                    .Where(test => state.ProjectsProvider.Component(test.Declaration).HasEqualCodeModule(component));
        }

        public static bool IsTestMethod(RubberduckParserState state, Declaration item)
        {
            return !state.AllUserDeclarations.Any(d =>
                       d.DeclarationType == DeclarationType.Parameter && Equals(d.ParentScopeDeclaration, item)) &&
                       item.Annotations.Any(pta => pta.Annotation is TestMethodAnnotation);
        }

        public static IEnumerable<Declaration> FindModuleInitializeMethods(QualifiedModuleName module, RubberduckParserState state)
        {
            return GetTestModuleProcedures(state)
                    .Where(m =>
                            m.QualifiedName.QualifiedModuleName == module &&
                            m.Annotations.Any(pta => pta.Annotation is ModuleInitializeAnnotation));
        }
        
        public static IEnumerable<Declaration> FindModuleCleanupMethods(QualifiedModuleName module, RubberduckParserState state)
        {
            return GetTestModuleProcedures(state)
                    .Where(m =>
                            m.QualifiedName.QualifiedModuleName == module &&
                            m.Annotations.Any(pta => pta.Annotation is ModuleCleanupAnnotation));
        }

        public static IEnumerable<Declaration> FindTestInitializeMethods(QualifiedModuleName module, RubberduckParserState state)
        {
            return GetTestModuleProcedures(state)
                    .Where(m =>
                            m.QualifiedName.QualifiedModuleName == module &&
                            m.Annotations.Any(pta => pta.Annotation is TestInitializeAnnotation));
        }

        public static IEnumerable<Declaration> FindTestCleanupMethods(QualifiedModuleName module, RubberduckParserState state)
        {
            return GetTestModuleProcedures(state)
                    .Where(m =>
                            m.QualifiedName.QualifiedModuleName == module &&
                            m.Annotations.Any(pta => pta.Annotation is TestCleanupAnnotation));
        }

        private static IEnumerable<Declaration> GetTestModuleProcedures(RubberduckParserState state)
        {
            var procedures = state.AllUserDeclarations.Where(item => item.DeclarationType == DeclarationType.Procedure);

            return procedures.Where(item =>
                        item.ParentDeclaration.DeclarationType == DeclarationType.ProceduralModule &&
                        item.ParentDeclaration.Annotations.Any(pta => pta.Annotation is TestModuleAnnotation));
        }

        public static IEnumerable<Declaration> GetTestModules(this RubberduckParserState state)
        {
            return state.AllUserDeclarations.Where(item =>
                        item.DeclarationType == DeclarationType.ProceduralModule &&
                        item.Annotations.Any(pta => pta.Annotation is TestModuleAnnotation));
        }
    }
}