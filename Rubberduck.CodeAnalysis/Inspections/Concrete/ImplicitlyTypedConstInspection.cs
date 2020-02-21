using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    public sealed class ImplicitlyTypedConstInspection : InspectionBase
    {
        public ImplicitlyTypedConstInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var declarationFinder = DeclarationFinderProvider.DeclarationFinder;

            var implicitlyTypedConsts = declarationFinder.UserDeclarations(DeclarationType.Constant)
                .Where(declaration => !declaration.IsTypeSpecified);

            return implicitlyTypedConsts.Select(Result);
        }

        private IInspectionResult Result(Declaration declaration)
        {
            var description = string.Format(InspectionResults.ImplicitlyTypedConstInspection, declaration.IdentifierName);

            return new DeclarationInspectionResult(
                this,
                description,
                declaration);
        }
    }
}
