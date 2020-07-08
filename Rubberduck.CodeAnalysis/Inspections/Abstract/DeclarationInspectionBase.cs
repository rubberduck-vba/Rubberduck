using System.Collections.Generic;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Results;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.Inspections.Abstract
{
    internal abstract class DeclarationInspectionBase : DeclarationInspectionBaseBase
    {
        protected DeclarationInspectionBase(IDeclarationFinderProvider declarationFinderProvider, params DeclarationType[] relevantDeclarationTypes)
            : base(declarationFinderProvider, relevantDeclarationTypes)
        {}

        protected DeclarationInspectionBase(IDeclarationFinderProvider declarationFinderProvider, DeclarationType[] relevantDeclarationTypes, DeclarationType[] excludeDeclarationTypes)
            : base(declarationFinderProvider, relevantDeclarationTypes, excludeDeclarationTypes)
        {}

        protected abstract bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder);
        protected abstract string ResultDescription(Declaration declaration);

        protected virtual ICollection<string> DisabledQuickFixes(Declaration declaration) => new List<string>();

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module, DeclarationFinder finder)
        {
            var objectionableDeclarations = RelevantDeclarationsInModule(module, finder)
                .Where(declaration => IsResultDeclaration(declaration, finder));

            return objectionableDeclarations
                .Select(InspectionResult)
                .ToList();
        }

        protected virtual IInspectionResult InspectionResult(Declaration declaration)
        {
            return new DeclarationInspectionResult(
                this,
                ResultDescription(declaration),
                declaration,
                disabledQuickFixes: DisabledQuickFixes(declaration));
        }
    }

    internal abstract class DeclarationInspectionBase<T> : DeclarationInspectionBaseBase
    {
        protected DeclarationInspectionBase(IDeclarationFinderProvider declarationFinderProvider, params DeclarationType[] relevantDeclarationTypes)
            : base(declarationFinderProvider, relevantDeclarationTypes)
        {}

        protected DeclarationInspectionBase(IDeclarationFinderProvider declarationFinderProvider , DeclarationType[] relevantDeclarationTypes, DeclarationType[] excludeDeclarationTypes)
            : base(declarationFinderProvider, relevantDeclarationTypes, excludeDeclarationTypes)
        {}

        protected abstract (bool isResult, T properties) IsResultDeclarationWithAdditionalProperties(Declaration declaration, DeclarationFinder finder);
        protected abstract string ResultDescription(Declaration declaration, T properties);

        protected virtual ICollection<string> DisabledQuickFixes(Declaration declaration, T properties) => new List<string>();

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module, DeclarationFinder finder)
        {
            var objectionableDeclarationsWithAdditionalProperties = RelevantDeclarationsInModule(module, finder)
                    .Select(declaration => DeclarationWithResultProperties(declaration, finder))
                    .Where(result => result.HasValue)
                    .Select(result => result.Value);

            return objectionableDeclarationsWithAdditionalProperties
                .Select(tpl => InspectionResult(tpl.declaration, tpl.properties))
                .ToList();
        }

        private (Declaration declaration, T properties)? DeclarationWithResultProperties(Declaration declaration, DeclarationFinder finder)
        {
            var (isResult, properties) = IsResultDeclarationWithAdditionalProperties(declaration, finder);
            return isResult
                ? (declaration, properties)
                : ((Declaration declaration, T properties)?) null;
        }

        protected virtual IInspectionResult InspectionResult(Declaration declaration, T properties)
        {
            return new DeclarationInspectionResult<T>(
                this,
                ResultDescription(declaration, properties),
                declaration,
                properties: properties,
                disabledQuickFixes: DisabledQuickFixes(declaration, properties));
        }
    }
}