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
    public abstract class IdentifierReferenceInspectionBase : InspectionBase
    {
        protected IdentifierReferenceInspectionBase(RubberduckParserState state)
            : base(state)
        {}

        protected abstract bool IsResultReference(IdentifierReference reference, DeclarationFinder finder);
        protected abstract string ResultDescription(IdentifierReference reference);

        protected virtual ICollection<string> DisabledQuickFixes(IdentifierReference reference) => new List<string>();

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var finder = DeclarationFinderProvider.DeclarationFinder;

            return finder.UserDeclarations(DeclarationType.Module)
                .Where(declaration => declaration != null)
                .SelectMany(declaration => DoGetInspectionResults(declaration.QualifiedModuleName, finder))
                .ToList();
        }

        protected IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module, DeclarationFinder finder)
        {
            var objectionableReferences = ReferencesInModule(module, finder)
                .Where(reference => IsResultReference(reference, finder));

            return objectionableReferences
                .Select(reference => InspectionResult(reference, DeclarationFinderProvider))
                .ToList();
        }

        protected IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module)
        {
            var finder = DeclarationFinderProvider.DeclarationFinder;
            return DoGetInspectionResults(module, finder);
        }

        protected virtual IEnumerable<IdentifierReference> ReferencesInModule(QualifiedModuleName module, DeclarationFinder finder)
        {
            return finder.IdentifierReferences(module);
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

    public abstract class IdentifierReferenceInspectionBase<T> : InspectionBase
    {
        protected IdentifierReferenceInspectionBase(RubberduckParserState state)
            : base(state)
        { }

        protected abstract (bool isResult, T properties) IsResultReferenceWithAdditionalProperties(IdentifierReference reference, DeclarationFinder finder);
        protected abstract string ResultDescription(IdentifierReference reference, T properties);

        protected virtual ICollection<string> DisabledQuickFixes(IdentifierReference reference, T properties) => new List<string>();

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var finder = DeclarationFinderProvider.DeclarationFinder;

            var results = new List<IInspectionResult>();
            foreach (var moduleDeclaration in State.DeclarationFinder.UserDeclarations(DeclarationType.Module))
            {
                if (moduleDeclaration == null)
                {
                    continue;
                }

                var module = moduleDeclaration.QualifiedModuleName;
                results.AddRange(DoGetInspectionResults(module, finder));
            }

            return results;
        }

        protected IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module, DeclarationFinder finder)
        {
            var objectionableReferencesWithProperties = ReferencesInModule(module, finder)
                .Select(reference => ReferenceWithResultProperties(reference, finder))
                .Where(result => result.HasValue)
                .Select(result => result.Value);

            return objectionableReferencesWithProperties
                .Select(tpl => InspectionResult(tpl.reference, DeclarationFinderProvider, tpl.properties))
                .ToList();
        }

        private (IdentifierReference reference, T properties)? ReferenceWithResultProperties(IdentifierReference reference, DeclarationFinder finder)
        {
            var (isResult, properties) = IsResultReferenceWithAdditionalProperties(reference, finder);
            return isResult
                ? (reference, properties)
                : ((IdentifierReference reference, T properties)?)null;
        }

        protected IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module)
        {
            var finder = DeclarationFinderProvider.DeclarationFinder;
            return DoGetInspectionResults(module, finder);
        }

        protected virtual IEnumerable<IdentifierReference> ReferencesInModule(QualifiedModuleName module, DeclarationFinder finder)
        {
            return finder.IdentifierReferences(module);
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