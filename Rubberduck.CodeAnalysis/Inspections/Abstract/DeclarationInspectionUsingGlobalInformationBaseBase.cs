using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Abstract
{
    public abstract class DeclarationInspectionUsingGlobalInformationBaseBase<T> : InspectionBase
    {
        protected readonly DeclarationType[] RelevantDeclarationTypes;
        protected readonly DeclarationType[] ExcludeDeclarationTypes;

        protected DeclarationInspectionUsingGlobalInformationBaseBase(RubberduckParserState state, params DeclarationType[] relevantDeclarationTypes)
            : base(state)
        {
            RelevantDeclarationTypes = relevantDeclarationTypes;
            ExcludeDeclarationTypes = new DeclarationType[0];
        }

        protected DeclarationInspectionUsingGlobalInformationBaseBase(RubberduckParserState state, DeclarationType[] relevantDeclarationTypes, DeclarationType[] excludeDeclarationTypes)
            : base(state)
        {
            RelevantDeclarationTypes = relevantDeclarationTypes;
            ExcludeDeclarationTypes = excludeDeclarationTypes;
        }

        protected abstract IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module, DeclarationFinder finder, T globalInformation);
        protected abstract T GlobalInformation(DeclarationFinder finder);

        protected virtual T GlobalInformation(QualifiedModuleName module, DeclarationFinder finder)
        {
            return GlobalInformation(finder);
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var finder = DeclarationFinderProvider.DeclarationFinder;
            var globalInformation = GlobalInformation(finder);

            var results = new List<IInspectionResult>();
            foreach (var moduleDeclaration in finder.UserDeclarations(DeclarationType.Module))
            {
                if (moduleDeclaration == null)
                {
                    continue;
                }

                var module = moduleDeclaration.QualifiedModuleName;
                results.AddRange(DoGetInspectionResults(module, finder, globalInformation));
            }

            foreach (var projectDeclaration in finder.UserDeclarations(DeclarationType.Project))
            {
                if (projectDeclaration == null)
                {
                    continue;
                }

                var module = projectDeclaration.QualifiedModuleName;
                results.AddRange(DoGetInspectionResults(module, finder, globalInformation));
            }

            return results;
        }

        protected virtual IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module)
        {
            var finder = DeclarationFinderProvider.DeclarationFinder;
            var globalInformation = GlobalInformation(module, finder);
            return DoGetInspectionResults(module, finder, globalInformation);
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