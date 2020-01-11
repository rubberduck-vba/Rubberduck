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
        protected IdentifierReferenceInspectionFromDeclarationsBase(RubberduckParserState state)
            : base(state)
        {}

        protected abstract IEnumerable<Declaration> ObjectionableDeclarations(DeclarationFinder finder);
        protected abstract string ResultDescription(IdentifierReference reference);

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var finder = DeclarationFinderProvider.DeclarationFinder;
            var objectionableReferences = ObjectionableReferences(finder);
            return objectionableReferences
                .Select(tpl => InspectionResult(tpl.reference, DeclarationFinderProvider, tpl.properties))
                .ToList();
        }

        protected virtual IEnumerable<(IdentifierReference reference, object properties)> ObjectionableReferences(DeclarationFinder finder)
        {
            var objectionableDeclarations = ObjectionableDeclarations(finder);
            return objectionableDeclarations
                .SelectMany(declaration => declaration.References)
                .Select(reference => (reference, IsResultReferenceWithAdditionalProperties(reference, finder)))
                .Where(tpl => tpl.Item2.isResult)
                .Select(tpl => (tpl.reference, tpl.Item2.properties));
        }

        protected virtual bool IsResultReference(IdentifierReference reference, DeclarationFinder finder) => true;

        protected virtual (bool isResult, object properties) IsResultReferenceWithAdditionalProperties(IdentifierReference reference, DeclarationFinder finder)
        {
            return (IsResultReference(reference, finder), null);
        }

        protected IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module)
        {
            var finder = DeclarationFinderProvider.DeclarationFinder;
            var objectionableReferences = ObjectionableReferences(finder);
            return objectionableReferences
                .Where(tpl => tpl.reference.QualifiedModuleName.Equals(module))
                .Select(tpl => InspectionResult(tpl.reference, DeclarationFinderProvider, tpl.properties))
                .ToList();
        }

        protected virtual IInspectionResult InspectionResult(IdentifierReference reference, IDeclarationFinderProvider declarationFinderProvider, dynamic properties = null)
        {
            return new IdentifierReferenceInspectionResult(
                this,
                ResultDescription(reference),
                declarationFinderProvider,
                reference,
                properties);
        }
    }
}