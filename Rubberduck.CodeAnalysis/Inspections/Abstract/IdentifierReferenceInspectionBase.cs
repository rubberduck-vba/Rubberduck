using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Abstract
{
    public abstract class IdentifierReferenceInspectionBase : InspectionBase
    {
        protected readonly IDeclarationFinderProvider DeclarationFinderProvider;

        public IdentifierReferenceInspectionBase(RubberduckParserState state)
            : base(state)
        {
            DeclarationFinderProvider = state;
        }

        protected abstract bool IsResultReference(IdentifierReference reference);
        protected abstract string ResultDescription(IdentifierReference reference);

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
            var objectionableReferences = ReferencesInModule(module)
                .Where(IsResultReference);

            return objectionableReferences
                .Select(reference => InspectionResult(reference, DeclarationFinderProvider))
                .ToList();
        }

        protected virtual IEnumerable<IdentifierReference> ReferencesInModule(QualifiedModuleName module)
        {
            return DeclarationFinderProvider.DeclarationFinder.IdentifierReferences(module);
        }

        protected virtual IInspectionResult InspectionResult(IdentifierReference reference, IDeclarationFinderProvider declarationFinderProvider)
        {
            return new IdentifierReferenceInspectionResult(
                this,
                ResultDescription(reference),
                declarationFinderProvider,
                reference);
        }
    }
}