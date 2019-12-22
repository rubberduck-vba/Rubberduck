using System.Linq;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class EmptyForEachInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void EmptyForEachBlock_InspectionName()
        {
            var inspection = new EmptyForEachBlockInspection(null);

            Assert.AreEqual(nameof(EmptyForEachBlockInspection), inspection.Name);
        }

        [Test]
        [Category("Inspections")]
        public void EmptyForEachBlock_DoesNotFiresOnImplementedLoopBlocks()
        {
            const string inputCode =
                @"Sub Foo(results As Collection)
    For Each var in results
        Msgbox Cstr(var)
    next var
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void EmptyForLoopBlock_FiresOnEmptyLoopBlocks()
        {
            const string inputCode =
                @"Sub Foo(results As Collection)
    For Each var in results
        'Msgbox Cstr(var)
    next var
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new EmptyForEachBlockInspection(state);
        }
    }
}
