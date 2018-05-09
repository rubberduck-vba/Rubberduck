using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class UseSetKeywordForObjectAssignmentQuickFixTests
    {
        [Test]
        [Category("QuickFixes")]
        public void ObjectVariableNotSet_ForFunctionAssignment_ReturnsResult()
        {
            var expectedResultCount = 2;
            var input =
                @"
Private Function CombineRanges(ByVal source As Range, ByVal toCombine As Range) As Range
    If source Is Nothing Then
        CombineRanges = toCombine 'no inspection result (but there should be one!)
    Else
        CombineRanges = Union(source, toCombine) 'no inspection result (but there should be one!)
    End If
End Function";
            var expectedCode =
                @"
Private Function CombineRanges(ByVal source As Range, ByVal toCombine As Range) As Range
    If source Is Nothing Then
        Set CombineRanges = toCombine 'no inspection result (but there should be one!)
    Else
        Set CombineRanges = Union(source, toCombine) 'no inspection result (but there should be one!)
    End If
End Function";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(input, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObjectVariableNotSetInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None).ToList();

                Assert.AreEqual(expectedResultCount, inspectionResults.Count);
                var fix = new UseSetKeywordForObjectAssignmentQuickFix(state);
                foreach (var result in inspectionResults)
                {
                    fix.Fix(result);
                }

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ObjectVariableNotSet_ForPropertyGetAssignment_ReturnsResults()
        {
            var expectedResultCount = 1;
            var input = @"
Private m_example As MyObject
Public Property Get Example() As MyObject
    Example = m_example
End Property
";
            var expectedCode =
                @"
Private m_example As MyObject
Public Property Get Example() As MyObject
    Set Example = m_example
End Property
";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(input, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObjectVariableNotSetInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None).ToList();

                Assert.AreEqual(expectedResultCount, inspectionResults.Count);
                var fix = new UseSetKeywordForObjectAssignmentQuickFix(state);
                foreach (var result in inspectionResults)
                {
                    fix.Fix(result);
                }

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

    }
}
