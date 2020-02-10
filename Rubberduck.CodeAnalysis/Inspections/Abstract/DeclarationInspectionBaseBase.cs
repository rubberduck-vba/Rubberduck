using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Abstract
{
    public abstract class DeclarationInspectionBaseBase : InspectionBase
    {
        protected readonly DeclarationType[] RelevantDeclarationTypes;
        protected readonly DeclarationType[] ExcludeDeclarationTypes;

        protected DeclarationInspectionBaseBase(RubberduckParserState state, params DeclarationType[] relevantDeclarationTypes)
            : base(state)
        {
            RelevantDeclarationTypes = relevantDeclarationTypes;
            ExcludeDeclarationTypes = new DeclarationType[0];
        }

        protected DeclarationInspectionBaseBase(RubberduckParserState state, DeclarationType[] relevantDeclarationTypes, DeclarationType[] excludeDeclarationTypes)
            : base(state)
        {
            RelevantDeclarationTypes = relevantDeclarationTypes;
            ExcludeDeclarationTypes = excludeDeclarationTypes;
        }

        protected abstract IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module, DeclarationFinder finder);

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

        protected virtual IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module)
        {
            var finder = DeclarationFinderProvider.DeclarationFinder;
            return DoGetInspectionResults(module, finder);
        }

        protected virtual IEnumerable<Declaration> RelevantDeclarationsInModule(QualifiedModuleName module, DeclarationFinder finder)
        {
            var potentiallyRelevantDeclarations = RelevantDeclarationTypes.Length == 0
                ? finder.Members(module)
                : RelevantDeclarationTypes
                    .SelectMany(declarationType => finder.Members(module, declarationType))
                    .Distinct();
            return potentiallyRelevantDeclarations
                .Where(declaration => !ExcludeDeclarationTypes.Contains(declaration.DeclarationType));
        }
    }
}