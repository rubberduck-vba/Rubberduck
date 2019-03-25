using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.Inspections.Inspections.Concrete.ThunderCode
{
    public class NonBreakingSpaceIdentifierInspection : InspectionBase
    {
        private const string Nbsp = "\u00A0";

        public NonBreakingSpaceIdentifierInspection(RubberduckParserState state) : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return State.DeclarationFinder.AllUserDeclarations
                .Where(d => d.IdentifierName.Contains(Nbsp))
                .Select(d => new DeclarationInspectionResult(
                    this, 
                    InspectionResults.NonBreakingSpaceIdentifierInspection.
                        ThunderCodeFormat(d.IdentifierName), 
                    d));
        }
    }
}
