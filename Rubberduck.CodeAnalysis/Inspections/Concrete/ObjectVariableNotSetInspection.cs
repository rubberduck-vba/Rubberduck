using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class ObjectVariableNotSetInspection : InspectionBase
    {
        public ObjectVariableNotSetInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {

            return InterestingReferences().Select(reference =>
                new IdentifierReferenceInspectionResult(this,
                    string.Format(InspectionResults.ObjectVariableNotSetInspection, reference.Declaration.IdentifierName),
                    State, reference));
        }

        private IEnumerable<IdentifierReference> InterestingReferences()
        {
            var result = new List<IdentifierReference>();
            foreach (var moduleReferences in State.DeclarationFinder.IdentifierReferences())
            {
                var module = State.DeclarationFinder.ModuleDeclaration(moduleReferences.Key);
                if (module == null || !module.IsUserDefined || IsIgnoringInspectionResultFor(module, AnnotationName))
                {
                    // module isn't user code (?), or this inspection is ignored at module-level
                    continue;
                }

                result.AddRange(moduleReferences.Value.Where(reference => !reference.IsSetAssignment
                    && VariableRequiresSetAssignmentEvaluator.RequiresSetAssignment(reference, State)));
            }

            return result.Where(reference => !IsIgnoringInspectionResultFor(reference, AnnotationName));
        }
    }
}
