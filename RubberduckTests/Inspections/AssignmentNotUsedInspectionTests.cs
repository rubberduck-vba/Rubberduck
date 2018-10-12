using NUnit.Framework;
using Rubberduck.Inspections.CodePathAnalysis;
using Rubberduck.Inspections.Concrete;
using RubberduckTests.Mocks;
using System.Linq;
using System.Threading;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class AssignmentNotUsedInspectionTests
    {
        [Test]
        [Category("Inspections")]
        public void MarksSequentialAssignments()
        {
            const string inputcode =
                @"Sub Foo()
    Dim i As Integer
    i = 9
    i = 8
    Bar i
End Sub

Sub Bar(ByVal i As Integer)
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputcode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new AssignmentNotUsedInspection(state, new Walker());
                var inspector = InspectionsHelper.GetInspector(inspection);
                var results = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, results.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void MarksLastAssignmentInDeclarationBlock()
        {
            const string inputcode =
                @"Sub Foo()
    Dim i As Integer
    i = 9
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputcode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new AssignmentNotUsedInspection(state, new Walker());
                var inspector = InspectionsHelper.GetInspector(inspection);
                var results = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, results.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        // Note: both assignments in the if/else can be marked in the future;
        // I just want feedback before I start mucking around that deep.
        public void DoesNotMarkLastAssignmentInNonDeclarationBlock()
        {
            const string inputcode =
                @"Sub Foo()
    Dim i As Integer
    i = 0
    If i = 9 Then
        i = 8
    Else
        i = 8
    End If
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputcode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new AssignmentNotUsedInspection(state, new Walker());
                var inspector = InspectionsHelper.GetInspector(inspection);
                var results = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(0, results.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void DoesNotMarkAssignmentWithReferenceAfter()
        {
            const string inputcode =
                @"Sub Foo()
    Dim i As Integer
    i = 9
    Bar i
End Sub

Sub Bar(ByVal i As Integer)
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputcode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new AssignmentNotUsedInspection(state, new Walker());
                var inspector = InspectionsHelper.GetInspector(inspection);
                var results = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(0, results.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void DoesNotMarkAssignment_UsedInForNext()
        {
            const string inputcode =
                @"Sub foo()
    Dim i As Integer
    i = 1
    For counter = i To 2
    Next
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputcode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new AssignmentNotUsedInspection(state, new Walker());
                var inspector = InspectionsHelper.GetInspector(inspection);
                var results = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(0, results.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void DoesNotMarkAssignment_UsedInWhileWend()
        {
            const string inputcode =
                @"Sub foo()
    Dim i As Integer
    i = 0

    While i < 10
        i = i + 1
    Wend
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputcode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new AssignmentNotUsedInspection(state, new Walker());
                var inspector = InspectionsHelper.GetInspector(inspection);
                var results = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(0, results.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void DoesNotMarkAssignment_UsedInDoWhile()
        {
            const string inputcode =
                @"Sub foo()
    Dim i As Integer
    i = 0
    Do While i < 10
    Loop
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputcode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new AssignmentNotUsedInspection(state, new Walker());
                var inspector = InspectionsHelper.GetInspector(inspection);
                var results = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(0, results.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void DoesNotMarkAssignment_UsedInSelectCase()
        {
            const string inputcode =
                @"Sub foo()
    Dim i As Integer
    i = 0
    Select Case i
        Case 0
            i = 1
        Case 1
            i = 2
        Case 2
            i = 3
        Case Else
            i = -1
    End Select
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputcode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new AssignmentNotUsedInspection(state, new Walker());
                var inspector = InspectionsHelper.GetInspector(inspection);
                var results = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(0, results.Count());
            }
        }
    }
}
