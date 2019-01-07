using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    [RequiredLibrary("Excel")]
    public sealed class ImplicitActiveWorkbookReferenceInspection : InspectionBase
    {
        public ImplicitActiveWorkbookReferenceInspection(RubberduckParserState state)
            : base(state) { }

        private static readonly string[] InterestingMembers =
        {
            "Worksheets", "Sheets", "Names"
        };

        private static readonly string[] InterestingClasses =
        {
            "_Global", "_Application", "Global", "Application"
        };

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var excel = State.DeclarationFinder.Projects.SingleOrDefault(item => !item.IsUserDefined && item.IdentifierName == "Excel");
            if (excel == null)
            {
                return Enumerable.Empty<IInspectionResult>();
            }

            var targetProperties = BuiltInDeclarations
                .OfType<PropertyGetDeclaration>()
                .Where(x => InterestingMembers.Contains(x.IdentifierName) && InterestingClasses.Contains(x.ParentDeclaration?.IdentifierName))
                .ToList();

            var members = targetProperties.SelectMany(item =>
                item.References.Where(reference => !IsIgnoringInspectionResultFor(reference, AnnotationName)));

            return members.Select(issue => new IdentifierReferenceInspectionResult(this,
                                                                string.Format(InspectionResults.ImplicitActiveWorkbookReferenceInspection, issue.Context.GetText()),
                                                                State,
                                                                issue));
        }
    }
}
