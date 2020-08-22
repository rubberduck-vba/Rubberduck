using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.Inspections.Abstract
{
    /// <summary>
    /// This is a base class for the other declaration inspection base classes. It should not be implemented directly by concrete inspections.
    /// </summary>
    internal abstract class DeclarationInspectionBaseBase : InspectionBase
    {
        private readonly DeclarationType[] _relevantDeclarationTypes;
        private readonly DeclarationType[] _excludeDeclarationTypes;

        protected DeclarationInspectionBaseBase(IDeclarationFinderProvider declarationFinderProvider, params DeclarationType[] relevantDeclarationTypes)
            : base(declarationFinderProvider)
        {
            _relevantDeclarationTypes = relevantDeclarationTypes;
            _excludeDeclarationTypes = new DeclarationType[0];
        }

        protected DeclarationInspectionBaseBase(IDeclarationFinderProvider declarationFinderProvider, DeclarationType[] relevantDeclarationTypes, DeclarationType[] excludeDeclarationTypes)
            : base(declarationFinderProvider)
        {
            _relevantDeclarationTypes = relevantDeclarationTypes;
            _excludeDeclarationTypes = excludeDeclarationTypes;
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults(DeclarationFinder finder)
        {
            return finder.UserDeclarations(DeclarationType.Module)
                .Concat(finder.UserDeclarations(DeclarationType.Project))
                .Where(declaration => declaration != null)
                .SelectMany(declaration => DoGetInspectionResults(declaration.QualifiedModuleName, finder))
                .ToList();
        }

        protected virtual IEnumerable<Declaration> RelevantDeclarationsInModule(QualifiedModuleName module, DeclarationFinder finder)
        {
            var potentiallyRelevantDeclarations = _relevantDeclarationTypes.Length == 0
                ? finder.Members(module)
                : _relevantDeclarationTypes
                    .SelectMany(declarationType => finder.Members(module, declarationType))
                    .Distinct();
            return potentiallyRelevantDeclarations
                .Where(declaration => !_excludeDeclarationTypes.Contains(declaration.DeclarationType));
        }
    }
}