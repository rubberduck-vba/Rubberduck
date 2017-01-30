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
        private readonly RubberduckParserState _state;

        public DebugDeclarations(RubberduckParserState state)
        {
            _state = state;
        }

        public IReadOnlyList<Declaration> Load()
        {
            var finder = _state.DeclarationFinder;;

            if (WeHaveAlreadyLoadedTheDeclarationsBefore(finder))
            {
                return new List<Declaration>();
            }

            var vba = finder.FindProject("VBA");
            if (vba == null)
            {
                // If the VBA project is null, we haven't loaded any COM references;
                // we're in a unit test and the mock project didn't setup any references.
                return new List<Declaration>();
            }

            return LoadDebugDeclarations(vba);
        }

        private static bool WeHaveAlreadyLoadedTheDeclarationsBefore(DeclarationFinder finder)
        {
            return ThereIsAGlobalBuiltInErrVariableDeclaration(finder);
        }

            private static bool ThereIsAGlobalBuiltInErrVariableDeclaration(DeclarationFinder finder) 
            {
                return finder.MatchName(Grammar.Tokens.Err).Any(declaration => declaration.IsBuiltIn
                                                                        && declaration.DeclarationType == DeclarationType.Variable
                                                                        && declaration.Accessibility == Accessibility.Global);
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
                    true,
                    new List<IAnnotation>(),
                    new Attributes());
}
                
                private static QualifiedModuleName DebugModuleName(Declaration parentProject)
                {
                    return new QualifiedModuleName(
                        parentProject.QualifiedName.QualifiedModuleName.ProjectName,
                        parentProject.QualifiedName.QualifiedModuleName.ProjectPath,
                        "DebugClass");
                }


            private static ClassModuleDeclaration DebugClassDeclaration(Declaration parentProject)
            {
                return new ClassModuleDeclaration(
                    new QualifiedMemberName(DebugClassName(parentProject), "DebugClass"), 
                    parentProject, 
                    "DebugClass", 
                    true, 
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
                    null);
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
                    true, 
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
                    true, 
                    null, 
                    new Attributes());
            }

    }
}
