using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Resources;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class EmptyForEachInspectionTests
    {
        [Test]
        [Category("Inspections")]
        public void EmptyForEachBlock_InspectionType()
        {
            var inspection = new EmptyForEachBlockInspection(null);
            var expectedInspection = CodeInspectionType.MaintainabilityAndReadabilityIssues;

            Assert.AreEqual(expectedInspection, inspection.InspectionType);
        }

        [Test]
        [Category("Inspections")]
        public void EmptyForEachBlock_InspectionName()
        {
            const string expectedName = nameof(EmptyForEachBlockInspection);
            var inspection = new EmptyForEachBlockInspection(null);

            Assert.AreEqual(expectedName, inspection.Name);
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
            CheckActualEmptyBlockCountEqualsExpected(inputCode, 0);
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
            CheckActualEmptyBlockCountEqualsExpected(inputCode, 1);
        }

        private void CheckActualEmptyBlockCountEqualsExpected(string inputCode, int expectedCount)
        {
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyForEachBlockInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(expectedCount, actualResults.Count());
            }
        }
    }
}
