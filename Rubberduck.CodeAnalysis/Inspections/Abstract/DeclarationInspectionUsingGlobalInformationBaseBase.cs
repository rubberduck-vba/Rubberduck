using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.Inspections.Abstract
{
    internal abstract class DeclarationInspectionUsingGlobalInformationBaseBase<T> : InspectionBase
    {
        protected readonly DeclarationType[] RelevantDeclarationTypes;
        protected readonly DeclarationType[] ExcludeDeclarationTypes;

        protected DeclarationInspectionUsingGlobalInformationBaseBase(IDeclarationFinderProvider declarationFinderProvider, params DeclarationType[] relevantDeclarationTypes)
            : base(declarationFinderProvider)
        {
            RelevantDeclarationTypes = relevantDeclarationTypes;
            ExcludeDeclarationTypes = new DeclarationType[0];
        }

        protected DeclarationInspectionUsingGlobalInformationBaseBase(IDeclarationFinderProvider declarationFinderProvider, DeclarationType[] relevantDeclarationTypes, DeclarationType[] excludeDeclarationTypes)
            : base(declarationFinderProvider)
        {
            RelevantDeclarationTypes = relevantDeclarationTypes;
            ExcludeDeclarationTypes = excludeDeclarationTypes;
        }

        protected abstract IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module, DeclarationFinder finder, T globalInformation);
        protected abstract T GlobalInformation(DeclarationFinder finder);

        /// <summary>
        /// Can be overwritten to enhance performance if providing the global information for one module module only is cheaper than getting it for all modules. 
        /// </summary>
        protected virtual T GlobalInformation(QualifiedModuleName module, DeclarationFinder finder)
        {
            return GlobalInformation(finder);
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults(DeclarationFinder finder)
        {
            var globalInformation = GlobalInformation(finder);

            return finder.UserDeclarations(DeclarationType.Module)
                .Concat(finder.UserDeclarations(DeclarationType.Project))
                .Where(declaration => declaration != null)
                .SelectMany(declaration => DoGetInspectionResults(declaration.QualifiedModuleName, finder, globalInformation))
                .ToList();
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module, DeclarationFinder finder)
        {
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