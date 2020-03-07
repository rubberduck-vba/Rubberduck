using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class RedundantOptionInspectionTests : InspectionTestsBase
    {
        [TestCase("Option Base 0", 1)]
        [TestCase("Option Base 1", 0)]
        [Category("Inspections")]
        public void RedundantOptionInspection_VariousScenario(string inputCode, int expectedCount)
        {
            Assert.AreEqual(expectedCount, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Ignore("Inspection class used here is 'ObsoleteCalStatementInspection - fails using RedundantOptionInspection")]
        [Category("Inspections")]
        public void RedundantOptionInspection_Ignored_DoesNotReturnResult()
        {
            var inputCode =
                @"'@Ignore OptionBaseZero
Option Base 0";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteCallStatementInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new RedundantOptionInspection(null);

            Assert.AreEqual(nameof(RedundantOptionInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new RedundantOptionInspection(state);
        }
    }
}
