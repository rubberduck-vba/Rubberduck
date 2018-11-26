using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class UseSetKeywordForObjectAssignmentQuickFixTests : QuickFixTestBase
    {
        [Test]
        [Category("QuickFixes")]
        public void ObjectVariableNotSet_ForFunctionAssignment_ReturnsResult()
        {
            var inputCode =
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

            var actualCode = ApplyQuickFixToAllInspectionResults(inputCode, state => new ObjectVariableNotSetInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ObjectVariableNotSet_ForPropertyGetAssignment_ReturnsResults()
        {
            var inputCode = @"
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

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ObjectVariableNotSetInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new UseSetKeywordForObjectAssignmentQuickFix();
        }
    }
}
