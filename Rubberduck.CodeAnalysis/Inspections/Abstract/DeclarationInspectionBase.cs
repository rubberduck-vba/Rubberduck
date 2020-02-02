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
    public abstract class DeclarationInspectionBase : InspectionBase
    {
        protected readonly DeclarationType[] RelevantDeclarationTypes;

        protected DeclarationInspectionBase(RubberduckParserState state, params DeclarationType[] relevantDeclarationTypes)
            : base(state)
        {
            RelevantDeclarationTypes = relevantDeclarationTypes;
        }

        protected abstract bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder);
        protected abstract string ResultDescription(Declaration declaration);

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
            var objectionableDeclarations = RelevantDeclarationsInModule(module, finder)
                .Where(declaration => IsResultDeclaration(declaration, finder));

            return objectionableDeclarations
                .Select(InspectionResult)
                .ToList();
        }

        protected virtual IEnumerable<Declaration> RelevantDeclarationsInModule(QualifiedModuleName module, DeclarationFinder finder)
        {
            return RelevantDeclarationTypes
                .SelectMany(declarationType => DeclarationFinderProvider.DeclarationFinder.Members(module, declarationType))
                .Distinct();
        }

        protected virtual IInspectionResult InspectionResult(Declaration declaration)
        {
            return new DeclarationInspectionResult(
                this,
                ResultDescription(declaration),
                declaration);
        }
    }
}