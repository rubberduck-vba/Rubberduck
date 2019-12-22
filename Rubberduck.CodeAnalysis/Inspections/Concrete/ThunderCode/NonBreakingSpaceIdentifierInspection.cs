using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.Inspections.Inspections.Concrete.ThunderCode
{
    /// <summary hidden="true">
    /// A ThunderCode inspection that locates non-breaking spaces hidden in identifier names.
    /// </summary>
    /// <why>
    /// This inpection is flagging code we dubbed "ThunderCode", 
    /// code our friend Andrew Jackson would have written to confuse Rubberduck's parser and/or resolver. 
    /// This inspection may accidentally reveal non-breaking spaces in code copied and pasted from a website.
    /// </why>
    public class NonBreakingSpaceIdentifierInspection : InspectionBase
    {
        private const string Nbsp = "\u00A0";

        public NonBreakingSpaceIdentifierInspection(RubberduckParserState state) : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return State.DeclarationFinder.AllUserDeclarations
                .Where(d => d.IdentifierName.Contains(Nbsp))
                .Select(d => new DeclarationInspectionResult(
                    this, InspectionResults.NonBreakingSpaceIdentifierInspection.ThunderCodeFormat(d.IdentifierName), d));
        }
    }
}
