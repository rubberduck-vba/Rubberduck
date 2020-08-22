using System.Collections.Generic;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Results;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.Inspections.Abstract
{
    internal abstract class DeclarationInspectionUsingGlobalInformationBase<TGlobalInfo> : DeclarationInspectionUsingGlobalInformationBaseBase<TGlobalInfo>
    {
        protected DeclarationInspectionUsingGlobalInformationBase(IDeclarationFinderProvider declarationFinderProvider, params DeclarationType[] relevantDeclarationTypes)
            : base(declarationFinderProvider, relevantDeclarationTypes)
        {}

        protected DeclarationInspectionUsingGlobalInformationBase(IDeclarationFinderProvider declarationFinderProvider, DeclarationType[] relevantDeclarationTypes, DeclarationType[] excludeDeclarationTypes)
            : base(declarationFinderProvider, relevantDeclarationTypes, excludeDeclarationTypes)
        {}

        protected abstract bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder, TGlobalInfo globalInfo);
        protected abstract string ResultDescription(Declaration declaration);

        protected virtual ICollection<string> DisabledQuickFixes(Declaration declaration) => new List<string>();

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module, DeclarationFinder finder, TGlobalInfo globalInfo)
        {
            var objectionableDeclarations = RelevantDeclarationsInModule(module, finder)
                .Where(declaration => IsResultDeclaration(declaration, finder, globalInfo));

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

    internal abstract class DeclarationInspectionUsingGlobalInformationBase<TGlobalInfo,TProperties> : DeclarationInspectionUsingGlobalInformationBaseBase<TGlobalInfo>
    {
        protected DeclarationInspectionUsingGlobalInformationBase(IDeclarationFinderProvider declarationFinderProvider, params DeclarationType[] relevantDeclarationTypes)
            : base(declarationFinderProvider, relevantDeclarationTypes)
        {}

        protected DeclarationInspectionUsingGlobalInformationBase(IDeclarationFinderProvider declarationFinderProvider, DeclarationType[] relevantDeclarationTypes, DeclarationType[] excludeDeclarationTypes)
            : base(declarationFinderProvider, relevantDeclarationTypes, excludeDeclarationTypes)
        {}

        protected abstract (bool isResult, TProperties properties) IsResultDeclarationWithAdditionalProperties(Declaration declaration, DeclarationFinder finder, TGlobalInfo globalInformation);
        protected abstract string ResultDescription(Declaration declaration, TProperties properties);

        protected virtual ICollection<string> DisabledQuickFixes(Declaration declaration, TProperties properties) => new List<string>();

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module, DeclarationFinder finder, TGlobalInfo globalInformation)
        {
            var objectionableDeclarationsWithAdditionalProperties = RelevantDeclarationsInModule(module, finder)
                .Select(declaration => DeclarationWithResultProperties(declaration, finder, globalInformation))
                .Where(result => result.HasValue)
                .Select(result => result.Value);

            return objectionableDeclarationsWithAdditionalProperties
                .Select(tpl => InspectionResult(tpl.declaration, tpl.properties))
                .ToList();
        }

        private (Declaration declaration, TProperties properties)? DeclarationWithResultProperties(Declaration declaration, DeclarationFinder finder, TGlobalInfo globalInformation)
        {
            var (isResult, properties) = IsResultDeclarationWithAdditionalProperties(declaration, finder, globalInformation);
            return isResult
                ? (declaration, properties)
                : ((Declaration declaration, TProperties properties)?)null;
        }

        protected virtual IInspectionResult InspectionResult(Declaration declaration, TProperties properties)
        {
            return new DeclarationInspectionResult<TProperties>(
                this,
                ResultDescription(declaration, properties),
                declaration,
                properties: properties,
                disabledQuickFixes: DisabledQuickFixes(declaration, properties));
        }
    }
}