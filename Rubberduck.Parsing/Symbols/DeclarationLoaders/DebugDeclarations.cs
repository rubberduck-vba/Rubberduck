using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols.DeclarationLoaders
{
    public class DebugDeclarations : ICustomDeclarationLoader
    {
        public static Declaration DebugPrint;
        private readonly IDeclarationFinderProvider _finderProvider;

        public DebugDeclarations(IDeclarationFinderProvider finderProvider)
        {
            _finderProvider = finderProvider;
        }

        public IReadOnlyList<Declaration> Load()
        {
            var finder = _finderProvider.DeclarationFinder;

            var vba = finder.FindProject("VBA");
            if (vba == null)
            {
                // If the VBA project is null, we haven't loaded any COM references;
                // we're in a unit test and the mock project didn't setup any references.
                return new List<Declaration>();
            }

            if (WeHaveAlreadyLoadedTheDeclarationsBefore(finder, vba))
            {
                return new List<Declaration>();
            }

            return LoadDebugDeclarations(vba);
        }

        private static bool WeHaveAlreadyLoadedTheDeclarationsBefore(DeclarationFinder finder, Declaration vbaProject)
        {
            return ThereIsADebugModule(finder, vbaProject);
        }

        private static bool ThereIsADebugModule(DeclarationFinder finder, Declaration vbaProject) 
        {
            var debugModule = DebugModuleDeclaration(vbaProject);
            return finder.MatchName(debugModule.IdentifierName)
                            .Any(declaration => declaration.Equals(debugModule));
        }


        private List<Declaration> LoadDebugDeclarations(Declaration parentProject)
        {
            var debugModule = DebugModuleDeclaration(parentProject);
            var debugClass = DebugClassDeclaration(parentProject); 
            var debugObject = DebugObjectDeclaration(debugModule);
            var debugAssert = DebugAssertDeclaration(debugClass);
            var debugPrint = DebugPrintDeclaration(debugClass);

            // Debug.Print has the same special syntax as the print and write statement.
            // Because of that it is treated specially in the grammar and normally wouldn't be resolved.
            // Since we still want it to be resolved we make it easier for the resolver to access the debug print
            // declaration by exposing it in this way.
            DebugPrint = debugPrint;

            return new List<Declaration> { 
                debugModule,
                debugClass,
                debugObject,
                debugAssert,
                debugPrint
            };
        }


        private static ProceduralModuleDeclaration DebugModuleDeclaration(Declaration parentProject)
        {
            return new ProceduralModuleDeclaration(
                new QualifiedMemberName(DebugModuleName(parentProject), "DebugModule"),
                parentProject,
                "DebugModule",
                false,
                new List<IAnnotation>(),
                new Attributes());
}
                
            private static QualifiedModuleName DebugModuleName(Declaration parentProject)
            {
                return new QualifiedModuleName(
                    parentProject.QualifiedName.QualifiedModuleName.ProjectName,
                    parentProject.QualifiedName.QualifiedModuleName.ProjectPath,
                    "DebugModule");
            }


        private static ClassModuleDeclaration DebugClassDeclaration(Declaration parentProject)
        {
            return new ClassModuleDeclaration(
                new QualifiedMemberName(DebugClassName(parentProject), "DebugClass"), 
                parentProject, 
                "DebugClass", 
                false, 
                new List<IAnnotation>(), 
                new Attributes(), 
                true);
        }

        private static QualifiedModuleName DebugClassName(Declaration parentProject)
        {
            return new QualifiedModuleName(
                parentProject.QualifiedName.QualifiedModuleName.ProjectName,
                parentProject.QualifiedName.QualifiedModuleName.ProjectPath,
                "DebugClass");
        }

        private static Declaration DebugObjectDeclaration(ProceduralModuleDeclaration debugModule)
        {
            return new Declaration(
                new QualifiedMemberName(debugModule.QualifiedName.QualifiedModuleName, "Debug"), 
                debugModule, 
                "Global", 
                "DebugClass", 
                null, 
                true, 
                false, 
                Accessibility.Global, 
                DeclarationType.Variable, 
                false, 
                null,
                false,
                null,
                new Attributes());
        }

        private static SubroutineDeclaration DebugAssertDeclaration(ClassModuleDeclaration debugClass)
        {
            return new SubroutineDeclaration(
                new QualifiedMemberName(debugClass.QualifiedName.QualifiedModuleName, "Assert"), 
                debugClass, 
                debugClass, 
                null, 
                Accessibility.Global, 
                null, 
                Selection.Home, 
                false, 
                null, 
                new Attributes());
        }

        private static SubroutineDeclaration DebugPrintDeclaration(ClassModuleDeclaration debugClass)
        {
            return new SubroutineDeclaration(
                new QualifiedMemberName(debugClass.QualifiedName.QualifiedModuleName, "Print"), 
                debugClass, 
                debugClass, 
                null, 
                Accessibility.Global, 
                null, Selection.Home, 
                false, 
                null, 
                new Attributes());
        }
    }
}
