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

        protected virtual ICollection<string> DisabledQuickFixes(IdentifierReference reference) => new List<string>();

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var finder = DeclarationFinderProvider.DeclarationFinder;
            var objectionableReferences = ObjectionableReferences(finder);
            var resultReferences = ResultReferences(objectionableReferences, finder);
            return resultReferences
                .Select(reference => InspectionResult(reference, DeclarationFinderProvider))
                .ToList();
        }

        private IEnumerable<IdentifierReference> ResultReferences(IEnumerable<IdentifierReference> potentialResultReferences, DeclarationFinder finder)
        {
            return potentialResultReferences
                .Where(reference => IsResultReference(reference, finder));
        }

        protected virtual IEnumerable<IdentifierReference> ObjectionableReferences(DeclarationFinder finder)
        {
            var objectionableDeclarations = ObjectionableDeclarations(finder);
            return objectionableDeclarations
                .SelectMany(declaration => declaration.References);
        }

        protected virtual bool IsResultReference(IdentifierReference reference, DeclarationFinder finder) => true;

        protected IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module)
        {
            var finder = DeclarationFinderProvider.DeclarationFinder;
            var objectionableReferences = ObjectionableReferences(finder)
                .Where(reference => reference.QualifiedModuleName.Equals(module));
            var resultReferences = ResultReferences(objectionableReferences, finder);
            return resultReferences
                .Select(reference => InspectionResult(reference, DeclarationFinderProvider))
                .ToList();
        }

        protected virtual IInspectionResult InspectionResult(IdentifierReference reference, IDeclarationFinderProvider declarationFinderProvider)
        {
            return new IdentifierReferenceInspectionResult(
                this,
                ResultDescription(reference),
                declarationFinderProvider,
                reference,
                DisabledQuickFixes(reference));
        }
    }

    public abstract class IdentifierReferenceInspectionFromDeclarationsBase<T> : InspectionBase
    {
        protected IdentifierReferenceInspectionFromDeclarationsBase(RubberduckParserState state)
            : base(state)
        { }

        protected abstract IEnumerable<Declaration> ObjectionableDeclarations(DeclarationFinder finder);
        protected abstract (bool isResult, T properties) IsResultReferenceWithAdditionalProperties(IdentifierReference reference, DeclarationFinder finder);
        protected abstract string ResultDescription(IdentifierReference reference, T properties);

        protected virtual ICollection<string> DisabledQuickFixes(IdentifierReference reference, T properties) => new List<string>();

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var finder = DeclarationFinderProvider.DeclarationFinder;
            var objectionableReferences = ObjectionableReferences(finder);
            var resultReferences = ResultReferences(objectionableReferences, finder);
            return resultReferences
                .Select(tpl => InspectionResult(tpl.reference, DeclarationFinderProvider, tpl.properties))
                .ToList();
        }

        private IEnumerable<(IdentifierReference reference, T properties)> ResultReferences(IEnumerable<IdentifierReference> potentialResultReferences, DeclarationFinder finder)
        {
            return potentialResultReferences
                .Select(reference => ReferenceWithResultProperties(reference, finder))
                .Where(result => result.HasValue)
                .Select(result => result.Value); ;
        }

        private (IdentifierReference reference, T properties)? ReferenceWithResultProperties(IdentifierReference reference, DeclarationFinder finder)
        {
            var (isResult, properties) = IsResultReferenceWithAdditionalProperties(reference, finder);
            return isResult
                ? (reference, properties)
                : ((IdentifierReference reference, T properties)?)null;
        }

        protected virtual IEnumerable<IdentifierReference> ObjectionableReferences(DeclarationFinder finder)
        {
            var objectionableDeclarations = ObjectionableDeclarations(finder);
            return objectionableDeclarations
                .SelectMany(declaration => declaration.References);
        }

        protected IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module)
        {
            var finder = DeclarationFinderProvider.DeclarationFinder;
            var objectionableReferences = ObjectionableReferences(finder)
                .Where(reference => reference.QualifiedModuleName.Equals(module));
            var resultReferences = ResultReferences(objectionableReferences, finder);
            return resultReferences
                .Select(tpl => InspectionResult(tpl.reference, DeclarationFinderProvider, tpl.properties))
                .ToList();
        }

        protected virtual IInspectionResult InspectionResult(IdentifierReference reference, IDeclarationFinderProvider declarationFinderProvider, T properties)
        {
            return new IdentifierReferenceInspectionResult<T>(
                this,
                ResultDescription(reference, properties),
                declarationFinderProvider,
                reference,
                properties,
                DisabledQuickFixes(reference, properties));
        }
    }
}