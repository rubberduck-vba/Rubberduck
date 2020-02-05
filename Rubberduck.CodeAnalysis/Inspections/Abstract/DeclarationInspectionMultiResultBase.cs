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
    public abstract class DeclarationInspectionMultiResultBase<T> : InspectionBase
    {
        protected readonly DeclarationType[] RelevantDeclarationTypes;
        protected readonly DeclarationType[] ExcludeDeclarationTypes;

        protected DeclarationInspectionMultiResultBase(RubberduckParserState state, params DeclarationType[] relevantDeclarationTypes)
            : base(state)
        {
            RelevantDeclarationTypes = relevantDeclarationTypes;
            ExcludeDeclarationTypes = new DeclarationType[0];
        }

        protected DeclarationInspectionMultiResultBase(RubberduckParserState state, DeclarationType[] relevantDeclarationTypes, DeclarationType[] excludeDeclarationTypes)
            : base(state)
        {
            RelevantDeclarationTypes = relevantDeclarationTypes;
            ExcludeDeclarationTypes = excludeDeclarationTypes;
        }

        protected abstract IEnumerable<T> ResultProperties(Declaration declaration, DeclarationFinder finder);
        protected abstract string ResultDescription(Declaration declaration, T properties);

        protected virtual ICollection<string> DisabledQuickFixes(Declaration declaration, T properties) => new List<string>();

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

        private IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module)
        {
            var finder = DeclarationFinderProvider.DeclarationFinder;
            return DoGetInspectionResults(module, finder);
        }

        private IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module, DeclarationFinder finder)
        {
            var objectionableDeclarationsWithAdditionalProperties = RelevantDeclarationsInModule(module, finder)
                    .SelectMany(declaration => ResultProperties(declaration, finder)
                                                .Select(properties => (declaration, properties)));

            return objectionableDeclarationsWithAdditionalProperties
                .Select(tpl => InspectionResult(tpl.declaration, tpl.properties))
                .ToList();
        }

        protected virtual IEnumerable<Declaration> RelevantDeclarationsInModule(QualifiedModuleName module, DeclarationFinder finder)
        {
            var potentiallyRelevantDeclarations = RelevantDeclarationTypes.Length == 0
                ? finder.AllUserDeclarations
                : RelevantDeclarationTypes
                    .SelectMany(declarationType => finder.Members(module, declarationType))
                    .Distinct();
            return potentiallyRelevantDeclarations
                .Where(declaration => ! ExcludeDeclarationTypes.Contains(declaration.DeclarationType));
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