using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class ObjectVariableNotSetInspection : InspectionBase
    {
        public ObjectVariableNotSetInspection(RubberduckParserState state)
            : base(state) {  }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {

            return InterestingReferences().Select(reference =>
                new IdentifierReferenceInspectionResult(this,
                    string.Format(InspectionsUI.ObjectVariableNotSetInspectionResultFormat, reference.Declaration.IdentifierName),
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

                foreach (var reference in moduleReferences.Value)
                {
                    if (!IsIgnoringInspectionResultFor(reference, AnnotationName) 
                        && VariableRequiresSetAssignmentEvaluator.NeedsSetKeywordAdded(reference, State))
                    {
                        result.Add(reference);
                    }
                }
            }

            return result;
        }
    }
}
