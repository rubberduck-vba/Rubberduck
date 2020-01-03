using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols.DeclarationLoaders
{
    public class DebugDeclarations : ICustomDeclarationLoader
    {
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
            var debugAssert = DebugAssertDeclaration(debugModule);
            var debugPrint = DebugPrintDeclaration(debugModule);

            return new List<Declaration> { 
                debugModule,
                debugAssert,
                debugPrint
            };
        }


        private static ProceduralModuleDeclaration DebugModuleDeclaration(Declaration parentProject)
        {
            return new ProceduralModuleDeclaration(
                new QualifiedMemberName(DebugModuleName(parentProject), Tokens.Debug),
                parentProject,
                "DebugModule",
                false,
                new List<IParseTreeAnnotation>(),
                new Attributes());
}
                
            private static QualifiedModuleName DebugModuleName(Declaration parentProject)
            {
                return new QualifiedModuleName(
                    parentProject.QualifiedName.QualifiedModuleName.ProjectName,
                    parentProject.QualifiedName.QualifiedModuleName.ProjectPath,
                    Tokens.Debug);
            }

        private static SubroutineDeclaration DebugAssertDeclaration(ProceduralModuleDeclaration debugModule)
        {
            return new SubroutineDeclaration(
                new QualifiedMemberName(debugModule.QualifiedName.QualifiedModuleName, "Assert"),
                debugModule,
                debugModule, 
                null, 
                Accessibility.Global, 
                null,
                null,
                Selection.Home, 
                false,
                new List<IParseTreeAnnotation>(), 
                new Attributes());
        }

        private static SubroutineDeclaration DebugPrintDeclaration(ProceduralModuleDeclaration debugModule)
        {
            return new SubroutineDeclaration(
                new QualifiedMemberName(debugModule.QualifiedName.QualifiedModuleName, "Print"),
                debugModule,
                debugModule, 
                null, 
                Accessibility.Global, 
                null, 
                null,
                Selection.Home, 
                false,
                new List<IParseTreeAnnotation>(), 
                new Attributes());
        }
    }
}
