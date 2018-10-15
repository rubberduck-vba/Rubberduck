using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.Inspections.Concrete
{
    public class IsMissingOnInappropriateArgumentInspection : IsMissingInspectionBase
    {
        public IsMissingOnInappropriateArgumentInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var results = new List<IInspectionResult>();

            foreach (var reference in IsMissingDeclarations.SelectMany(decl => decl.References.Where(candidate => !IsIgnoringInspectionResultFor(candidate, AnnotationName))))
            {
                var parameter = GetParameterForReference(reference);

                if (parameter == null || 
                    parameter.IsOptional 
                    && parameter.AsTypeName.Equals(Tokens.Variant) 
                    && string.IsNullOrEmpty(parameter.DefaultValue) 
                    && !parameter.IsArray)
                {
                    continue;                   
                }

                results.Add(new IdentifierReferenceInspectionResult(this, InspectionResults.IsMissingOnInappropriateArgumentInspection, State, reference, parameter));
            }

            return results;
        }
    }
}
