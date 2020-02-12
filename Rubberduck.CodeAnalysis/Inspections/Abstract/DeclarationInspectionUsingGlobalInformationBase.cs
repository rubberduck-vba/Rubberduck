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
    public abstract class DeclarationInspectionUsingGlobalInformationBase<TGlobalInfo> : DeclarationInspectionUsingGlobalInformationBaseBase<TGlobalInfo>
    {
        protected DeclarationInspectionUsingGlobalInformationBase(RubberduckParserState state, params DeclarationType[] relevantDeclarationTypes)
            : base(state, relevantDeclarationTypes)
        {}

        protected DeclarationInspectionUsingGlobalInformationBase(RubberduckParserState state, DeclarationType[] relevantDeclarationTypes, DeclarationType[] excludeDeclarationTypes)
            : base(state, relevantDeclarationTypes, excludeDeclarationTypes)
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

    public abstract class DeclarationInspectionUsingGlobalInformationBase<TGlobalInfo,TProperties> : DeclarationInspectionUsingGlobalInformationBaseBase<TGlobalInfo>
    {
        protected DeclarationInspectionUsingGlobalInformationBase(RubberduckParserState state, params DeclarationType[] relevantDeclarationTypes)
            : base(state, relevantDeclarationTypes)
        {}

        protected DeclarationInspectionUsingGlobalInformationBase(RubberduckParserState state, DeclarationType[] relevantDeclarationTypes, DeclarationType[] excludeDeclarationTypes)
            : base(state, relevantDeclarationTypes, excludeDeclarationTypes)
        {}

        protected abstract (bool isResult, TProperties properties) IsResultDeclarationWithAdditionalProperties(Declaration declaration, DeclarationFinder finder, TGlobalInfo globalInformation);
        protected abstract string ResultDescription(Declaration declaration, TProperties properties);

        protected virtual ICollection<string> DisabledQuickFixes(Declaration declaration, TProperties properties) => new List<string>();

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module, DeclarationFinder finder, TGlobalInfo globalInformation)
        {
            var objectionableDeclarationsWithAdditionalProperties = RelevantDeclarationsInModule(module, finder)
                    .Select(declaration => (declaration, IsResultDeclarationWithAdditionalProperties(declaration, finder, globalInformation)))
                    .Where(tpl => tpl.Item2.isResult)
                    .Select(tpl => (tpl.declaration, tpl.Item2.properties));

            return objectionableDeclarationsWithAdditionalProperties
                .Select(tpl => InspectionResult(tpl.declaration, tpl.properties))
                .ToList();
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