using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    public sealed class ImplicitlyTypedConstInspection : InspectionBase
    {
        public ImplicitlyTypedConstInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var declarationFinder = State.DeclarationFinder;

            var implicitlyTypedConsts = declarationFinder.AllDeclarations
                .Where(declaration => declaration.DeclarationType == DeclarationType.Constant
                    && declaration.AsTypeContext == null);

            return implicitlyTypedConsts.Select(Result);
        }

        private IInspectionResult Result(Declaration declaration)
        {
            return new IdentifierReferenceInspectionResult(
                this,
                declaration.DescriptionString,
                State,
                new IdentifierReference(
                    declaration.QualifiedModuleName,
                    declaration.ParentScopeDeclaration,
                    declaration.ParentDeclaration,
                    declaration.IdentifierName,
                    declaration.Selection,
                    declaration.Context,
                    declaration));
        }
    }
}
