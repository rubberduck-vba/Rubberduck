using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
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

        protected abstract bool IsResultDeclaration(Declaration declaration);
        protected abstract string ResultDescription(Declaration declaration);

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var results = new List<IInspectionResult>();
            foreach (var moduleDeclaration in State.DeclarationFinder.UserDeclarations(DeclarationType.Module))
            {
                if (moduleDeclaration == null)
                {
                    continue;
                }

                var module = moduleDeclaration.QualifiedModuleName;
                results.AddRange(DoGetInspectionResults(module));
            }

            return results;
        }

        private IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module)
        {
            var objectionableDeclarations = RelevantDeclarationsInModule(module)
                .Where(IsResultDeclaration);

            return objectionableDeclarations
                .Select(InspectionResult)
                .ToList();
        }

        protected virtual IEnumerable<Declaration> RelevantDeclarationsInModule(QualifiedModuleName module)
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