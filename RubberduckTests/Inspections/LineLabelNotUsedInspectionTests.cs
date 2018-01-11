using System.Linq;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Resources;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class LineLabelNotUsedInspectionTests
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
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new LineLabelNotUsedInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(5, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void LabelNotUsed_ReturnsResult()
        {
            const string inputCode =
                @"Sub Foo()
label1:
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new LineLabelNotUsedInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(1, inspectionResults.Count());
            }
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
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new LineLabelNotUsedInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(2, inspectionResults.Count());
            }
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
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new LineLabelNotUsedInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count());
            }
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
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new LineLabelNotUsedInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count());
            }
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
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new LineLabelNotUsedInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count());
            }
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
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new LineLabelNotUsedInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void LabelNotUsed_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"Sub Foo()
'@Ignore label1
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new LineLabelNotUsedInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void InspectionType()
        {
            var inspection = new LineLabelNotUsedInspection(null);
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "LineLabelNotUsedInspection";
            var inspection = new LineLabelNotUsedInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
