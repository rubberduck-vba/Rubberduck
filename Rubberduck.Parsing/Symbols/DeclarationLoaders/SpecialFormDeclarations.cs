using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.VBA.DeclarationCaching;

namespace Rubberduck.Parsing.Symbols.DeclarationLoaders
{
    public class SpecialFormDeclarations : ICustomDeclarationLoader
    {
        private readonly IDeclarationFinderProvider _finderProvider;

        public SpecialFormDeclarations(IDeclarationFinderProvider finderProvider)
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

            var informationModule = finder.FindStdModule("Information", vba, true);
            if (informationModule == null)
            {
                //This should not happen under normal circumstances.
                //Most probably, we are in a test that only added parts of the VBA project.
                return new List<Declaration>();
            }

            if (WeHaveAlreadyLoadedTheDeclarationsBefore(finder, informationModule))
            {
                return new List<Declaration>();
            }

            return LoadSpecialFormDeclarations(informationModule);
        }

        private static bool WeHaveAlreadyLoadedTheDeclarationsBefore(DeclarationFinder finder, Declaration informationModule)
        {
            return ThereIsAnLBoundFunctionDeclaration(finder, informationModule);
        }

        private static bool ThereIsAnLBoundFunctionDeclaration(DeclarationFinder finder, Declaration InformationModule)
        {
            var lBoundFunction = LBoundFunction(InformationModule);
            return finder.MatchName(lBoundFunction.IdentifierName)
                            .Any(declaration => declaration.Equals(lBoundFunction));
        }

        private List<Declaration> LoadSpecialFormDeclarations(Declaration parentModule)
        {
            Debug.Assert(parentModule != null);

            var lboundFunction = LBoundFunction(parentModule);
            var uboundFunction = UBoundFunction(parentModule);

            return new List<Declaration> {
                lboundFunction,
                uboundFunction
            };
        }

        private static FunctionDeclaration LBoundFunction(Declaration parentModule)
        {
            var lboundFunction = LBoundFunctionWithoutParameters(parentModule);
            lboundFunction.AddParameter(ArrayNameParameter(parentModule, lboundFunction));
            lboundFunction.AddParameter(DimensionParameter(parentModule, lboundFunction));
            return lboundFunction;
        }

        private static FunctionDeclaration LBoundFunctionWithoutParameters(Declaration parentModule)
        {
            return new FunctionDeclaration(
                new QualifiedMemberName(parentModule.QualifiedName.QualifiedModuleName, "LBound"),
                parentModule,
                parentModule,
                "Long",
                null,
                null,
                Accessibility.Public,
                null,
                null,
                Selection.Home,
                false,
                false,
                new List<IParseTreeAnnotation>(),
                new Attributes());
        }

        private static ParameterDeclaration ArrayNameParameter(Declaration parentModule, FunctionDeclaration parentFunction)
        {
            var arrayParam = new ParameterDeclaration(
                new QualifiedMemberName(parentModule.QualifiedName.QualifiedModuleName, "Arrayname"), 
                parentFunction, 
                "Variant", 
                null, 
                null, 
                false, 
                false, 
                true);
            return arrayParam;
        }

        private static ParameterDeclaration DimensionParameter(Declaration parentModule, FunctionDeclaration parentFunction)
        {
            var rankParam = new ParameterDeclaration(
                new QualifiedMemberName(parentModule.QualifiedName.QualifiedModuleName, "Dimension"), 
                parentFunction, 
                "Long", 
                null, 
                null, 
                true, 
                false);
            return rankParam;
        }

        private static FunctionDeclaration UBoundFunction(Declaration parentModule)
        {
            var uBoundFunction = UBoundFunctionWithoutParameters(parentModule);
            uBoundFunction.AddParameter(ArrayNameParameter(parentModule, uBoundFunction));
            uBoundFunction.AddParameter(DimensionParameter(parentModule, uBoundFunction));
            return uBoundFunction;
        }

        private static FunctionDeclaration UBoundFunctionWithoutParameters(Declaration parentModule)
        {
            return new FunctionDeclaration(
                new QualifiedMemberName(parentModule.QualifiedName.QualifiedModuleName, "UBound"),
                parentModule,
                parentModule,
                "Long",
                null,
                null,
                Accessibility.Public,
                null,
                null,
                Selection.Home,
                false,
                false,
                new List<IParseTreeAnnotation>(),
                new Attributes());
        }
    }
}
