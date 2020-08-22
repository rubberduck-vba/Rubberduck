using System.Collections.Generic;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Results;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.Inspections.Abstract
{
    internal abstract class DeclarationInspectionMultiResultBase<T> : DeclarationInspectionBaseBase
    {
        protected DeclarationInspectionMultiResultBase(IDeclarationFinderProvider declarationFinderProvider, params DeclarationType[] relevantDeclarationTypes)
            : base(declarationFinderProvider, relevantDeclarationTypes)
        {}

        protected DeclarationInspectionMultiResultBase(IDeclarationFinderProvider declarationFinderProvider, DeclarationType[] relevantDeclarationTypes, DeclarationType[] excludeDeclarationTypes)
            : base(declarationFinderProvider, relevantDeclarationTypes, excludeDeclarationTypes)
        {}

        protected abstract IEnumerable<T> ResultProperties(Declaration declaration, DeclarationFinder finder);
        protected abstract string ResultDescription(Declaration declaration, T properties);

        protected virtual ICollection<string> DisabledQuickFixes(Declaration declaration, T properties) => new List<string>();

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module, DeclarationFinder finder)
        {
            var objectionableDeclarationsWithAdditionalProperties = RelevantDeclarationsInModule(module, finder)
                    .SelectMany(declaration => ResultProperties(declaration, finder)
                                                .Select(properties => (declaration, properties)));

            return objectionableDeclarationsWithAdditionalProperties
                .Select(tpl => InspectionResult(tpl.declaration, tpl.properties))
                .ToList();
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