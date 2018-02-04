using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Resources;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class EmptyIfBlockInspectionTests
    {
        [Test]
        [Category("Inspections")]
        public void EmptyIfBlock_FiresOnEmptyIfBlock()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new EmptyIfBlockInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void EmptyIfBlock_FiresOnEmptyElseIfBlock()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
    ElseIf False Then
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyIfBlockInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(2, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void EmptyIfBlock_FiresOnEmptyIfBlock_ElseBlock()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
    Else
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyIfBlockInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void EmptyIfBlock_FiresOnEmptySingleLineIfStmt()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then Else Bar
End Sub

Sub Bar()
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyIfBlockInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void EmptyIfBlock_FiresOnEmptyIfBlock_HasNonEmptyElseBlock()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
    Else
        Dim d
        d = 0
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyIfBlockInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void EmptyIfBlock_FiresOnEmptyIfBlock_HasQuoteComment()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
        ' Im a comment
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyIfBlockInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void EmptyIfBlock_FiresOnEmptyIfBlock_HasRemComment()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
        Rem Im a comment
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyIfBlockInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void EmptyIfBlock_FiresOnEmptyIfBlock_HasVariableDeclaration()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
        Dim d
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyIfBlockInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void EmptyIfBlock_FiresOnEmptyIfBlock_HasConstDeclaration()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
        Const c = 0
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyIfBlockInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void EmptyIfBlock_FiresOnEmptyIfBlock_HasWhitespace()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then

    	
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyIfBlockInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void EmptyIfBlock_IfBlockHasExecutableStatement()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
        Dim d
        d = 0
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyIfBlockInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void EmptyIfBlock_SingleLineIfBlockHasExecutableStatement()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then Bar
End Sub

Sub Bar()
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyIfBlockInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void EmptyIfBlock_IfAndElseIfBlockHaveExecutableStatement()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
        Dim d
        d = 0
    ElseIf False Then
        Dim b
        b = 0
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyIfBlockInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void InspectionType()
        {
            var inspection = new EmptyIfBlockInspection(null);
            Assert.AreEqual(CodeInspectionType.MaintainabilityAndReadabilityIssues, inspection.InspectionType);
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = nameof(EmptyIfBlockInspection);
            var inspection = new EmptyIfBlockInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
