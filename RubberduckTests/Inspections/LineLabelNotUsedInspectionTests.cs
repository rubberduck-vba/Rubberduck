using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class LineLabelNotUsedInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void LabelNotUsed_EdgeCaseIssue3226()
        {
            const string inputCode = @"
Sub foo()

100           'line-number
200:          'line-number with instruction spearator
300 Beep      'Line-number and instruction
400: Beep     'Line-number with instruction separator and instruction

bar:         'line-label
buzz: Beep   'Line-label and instruction
50 fizz:     'line-number and line-label
10 foo: Beep 'Line number and line-label (that matches procedure name) and instruction
20 boo: Beep 'Line number and line-label and instruction

End Sub
";
            Assert.AreEqual(5, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void LabelNotUsed_ReturnsResult()
        {
            const string inputCode =
                @"Sub Foo()
label1:
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void LabelNotUsed_ReturnsResult_MultipleLabels()
        {
            const string inputCode =
                @"Sub Foo()
label1:
label2:
End Sub";
            Assert.AreEqual(2, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void GivenGoToStatement_LabelUsed_YieldsNoResult()
        {
            const string inputCode =
                @"Sub Foo()
label1:
    GoTo label1
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void GivenGoToStatement_GoToBeforeLabel_LabelUsed_YieldsNoResult()
        {
            const string inputCode =
                @"Sub Foo()
    GoTo label1
label1:
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void GivenMultipleGoToStatements_BothLabelsUsed_YieldsNoResult()
        {
            const string inputCode =
                @"Sub Foo()
label1:
    GoTo label1
label2:
    Goto label2
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void LabelNotUsed_ReturnsResult_WithUsedLabelThatDoesNotReturnResult()
        {
            const string inputCode =
                @"Sub Foo()
    GoTo Label1:

label2:

label1:
End Sub

Sub Goo(ByVal arg1 As Integer)
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void LabelNotUsed_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"Sub Foo()
'@Ignore label1
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new LineLabelNotUsedInspection(null);

            Assert.AreEqual(nameof(LineLabelNotUsedInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new LineLabelNotUsedInspection(state);
        }
    }
}
