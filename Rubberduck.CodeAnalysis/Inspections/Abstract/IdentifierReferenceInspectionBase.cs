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
    public abstract class IdentifierReferenceInspectionBase : InspectionBase
    {
        protected IdentifierReferenceInspectionBase(RubberduckParserState state)
            : base(state)
        {}

        protected abstract bool IsResultReference(IdentifierReference reference, DeclarationFinder finder);
        protected abstract string ResultDescription(IdentifierReference reference, dynamic properties = null);

        protected virtual (bool isResult, object properties) IsResultReferenceWithAdditionalProperties(IdentifierReference reference, DeclarationFinder finder)
        {
            return (IsResultReference(reference, finder), null);
        }

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

        protected IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module, DeclarationFinder finder)
        {
            var objectionableReferencesWithProperties = ReferencesInModule(module, finder)
                .Select(reference => (reference, IsResultReferenceWithAdditionalProperties(reference, finder)))
                .Where(tpl => tpl.Item2.isResult)
                .Select(tpl => (tpl.reference, tpl.Item2.properties));

            return objectionableReferencesWithProperties
                .Select(tpl => InspectionResult(tpl.reference, DeclarationFinderProvider, tpl.properties))
                .ToList();
        }

        protected IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module)
        {
            var finder = DeclarationFinderProvider.DeclarationFinder;
            return DoGetInspectionResults(module, finder);
        }

        protected virtual IEnumerable<IdentifierReference> ReferencesInModule(QualifiedModuleName module, DeclarationFinder finder)
        {
            return finder.IdentifierReferences(module);
        }

        protected virtual IInspectionResult InspectionResult(IdentifierReference reference, IDeclarationFinderProvider declarationFinderProvider, dynamic properties = null)
        {
            return new IdentifierReferenceInspectionResult(
                this,
                ResultDescription(reference, properties),
                declarationFinderProvider,
                reference,
                properties);
        }
    }
}