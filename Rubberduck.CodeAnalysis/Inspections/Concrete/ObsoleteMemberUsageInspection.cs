using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.Inspections.Inspections.Concrete
{
    public sealed class ObsoleteMemberUsageInspection : InspectionBase
    {
        public ObsoleteMemberUsageInspection(RubberduckParserState state) : base(state)
        {
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return State.AllUserDeclarations
                .Where(declaration => declaration.DeclarationType.HasFlag(DeclarationType.Member) &&
                                      declaration.Annotations.Any(annotation =>
                                          annotation.AnnotationType == AnnotationType.Obsolete))
                .SelectMany(declaration => declaration.References).Select(reference =>
                    new IdentifierReferenceInspectionResult(this,
                        string.Format(InspectionResults.ObsoleteMemberUsageInspection, reference.IdentifierName), State,
                        reference));
        }
    }
}
