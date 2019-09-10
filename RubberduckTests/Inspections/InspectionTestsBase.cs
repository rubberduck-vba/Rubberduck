using System.Collections.Generic;
using System.Threading;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    public abstract class InspectionTestsBase
    {
        protected abstract IInspection InspectionUnderTest(RubberduckParserState state);

        public IEnumerable<IInspectionResult> InspectionResults(IVBE vbe)
        {
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var inspection = InspectionUnderTest(state);
                return InspectionResults(inspection, state);
            }
        }

        private static IEnumerable<IInspectionResult> InspectionResults(IInspection inspection, RubberduckParserState state)
        {
            if (inspection is IParseTreeInspection)
            {
                var inspector = InspectionsHelper.GetInspector(inspection);
                return inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            }

            return inspection.GetInspectionResults(CancellationToken.None);
        }
    }
}