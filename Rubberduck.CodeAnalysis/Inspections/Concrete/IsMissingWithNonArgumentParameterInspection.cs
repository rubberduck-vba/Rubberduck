using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.Inspections.Inspections.Concrete
{
    public class IsMissingWithNonArgumentParameterInspection : IsMissingInspectionBase
    {
        public IsMissingWithNonArgumentParameterInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var results = new List<IInspectionResult>();

            foreach (var reference in IsMissingDeclarations.SelectMany(decl => decl.References.Where(candidate => !IsIgnoringInspectionResultFor(candidate, AnnotationName))))
            {
                var parameter = GetParameterForReference(reference);

                if (parameter != null)
                {
                    continue;
                }

                results.Add(new IdentifierReferenceInspectionResult(this, InspectionResults.IsMissingWithNonArgumentParameterInspection, State, reference));
            }

            return results;
        }
    }
}
