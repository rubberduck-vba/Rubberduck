using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Abstract
{
    public abstract class IdentifierReferenceInspectionFromDeclarationsBase : InspectionBase
    {
        protected readonly IDeclarationFinderProvider DeclarationFinderProvider;

        protected IdentifierReferenceInspectionFromDeclarationsBase(RubberduckParserState state)
            : base(state)
        {
            DeclarationFinderProvider = state;
        }

        protected abstract IEnumerable<Declaration> ObjectionableDeclarations(DeclarationFinder finder);
        protected abstract string ResultDescription(IdentifierReference reference);

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var finder = DeclarationFinderProvider.DeclarationFinder;
            var objectionableReferences = ObjectionableReferences(finder);
            return objectionableReferences
                .Select(reference => InspectionResult(reference, DeclarationFinderProvider))
                .ToList();
        }

        private IEnumerable<IdentifierReference> ObjectionableReferences(DeclarationFinder finder)
        {
            var objectionableDeclarations = ObjectionableDeclarations(finder);
            return objectionableDeclarations
                .SelectMany(declaration => declaration.References)
                .Where(IsResultReference);
        }

        protected virtual bool IsResultReference(IdentifierReference reference) => true;

        protected IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module)
        {
            var finder = DeclarationFinderProvider.DeclarationFinder;
            var objectionableReferences = ObjectionableReferences(finder);
            return objectionableReferences
                .Where(reference => reference.QualifiedModuleName.Equals(module))
                .Select(reference => InspectionResult(reference, DeclarationFinderProvider))
                .ToList();
        }

        protected virtual IInspectionResult InspectionResult(IdentifierReference reference, IDeclarationFinderProvider declarationFinderProvider)
        {
            return new IdentifierReferenceInspectionResult(
                this,
                ResultDescription(reference),
                declarationFinderProvider,
                reference);
        }
    }
}