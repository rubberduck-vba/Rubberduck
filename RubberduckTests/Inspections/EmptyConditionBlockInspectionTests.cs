using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RubberduckTests.Mocks;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Inspections.QuickFixes;

namespace RubberduckTests.Inspections
{
    
    [TestClass]
    public class EmptyConditionBlockInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new EmptyConditionBlockInspection(null);

            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = nameof(EmptyConditionBlockInspection);
            var inspection = new EmptyConditionBlockInspection(null);
            
            Assert.AreEqual(inspectionName, inspection.Name);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_FiresOnEmptyIfBlock()
        {
            const string inputCode =
@"Sub Foo()
    If True Then
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_FiresOnEmptyElseIfBlock()
        {
            const string inputCode =
@"Sub Foo()
    If True Then
    ElseIf False Then
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_FiresOnEmptyIfBlock_ElseBlock()
        {
            const string inputCode =
@"Sub Foo()
    If True Then
        Dim a as Long
        a = 0
    Else
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_FiresOnEmptySingleLineIfStmt()
        {
            const string inputCode =
@"Sub Foo()
    If True Then Else Bar
End Sub

Sub Bar()
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
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
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_FiresOnEmptyIfBlock_HasQuoteComment()
        {
            const string inputCode =
@"Sub Foo()
    If True Then
        ' Im a comment
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_FiresOnEmptyIfBlock_HasRemComment()
        {
            const string inputCode =
@"Sub Foo()
    If True Then
        Rem Im a comment
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_FiresOnEmptyIfBlock_HasVariableDeclaration()
        {
            const string inputCode =
@"Sub Foo()
    If True Then
        Dim d
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_FiresOnEmptyIfBlock_HasConstDeclaration()
        {
            const string inputCode =
@"Sub Foo()
    If True Then
        Const c = 0
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_FiresOnEmptyIfBlock_HasWhitespace()
        {
            const string inputCode =
@"Sub Foo()
    If True Then

    	
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
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
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_SingleLineIfBlockHasExecutableStatement()
        {
            const string inputCode =
@"Sub Foo()
    If True Then Bar
End Sub

Sub Bar()
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
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
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyElseBlock_FiresOnEmptyIfBlock()
        {
            const string inputcode =
@"Sub Foo()
    If True Then
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputcode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            const int expectedCount = 1;

            Assert.AreEqual(expectedCount, actualResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyElseBlock_HasNoContent()
        {
            const string inputcode =
@"Sub Foo()
    If True Then
        Dim a as Long
        a = 0
    Else
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputcode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            const int expectedCount = 1;

            Assert.AreEqual(expectedCount, actualResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyElseBlock_HasQuoteComment()
        {
            const string inputCode =
@"Sub Foo()
    If True Then
        Dim a as Long
        a = 0
    Else
        'Some Comment
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            const int expectedCount = 1;

            Assert.AreEqual(expectedCount, actualResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyElseBlock_HasRemComment()
        {
            const string inputCode =
@"Sub Foo()
    If True Then
        Dim a as Long
        a = 0
    Else
        Rem Some Comment
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            const int expectedCount = 1;

            Assert.AreEqual(expectedCount, actualResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyElseBlock_HasVariableDeclaration()
        {
            const string inputCode =
@"Sub Foo()
    If True Then
        Dim a as Long
        a = 0
    Else
        Dim d
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            const int expectedCount = 1;

            Assert.AreEqual(expectedCount, actualResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyElseBlock_HasConstDeclaration()
        {
            const string inputCode =
@"Sub Foo()
    If True Then
        Dim a as Long
        a = 0
    Else
        Const c = 0
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            const int expectedCount = 1;

            Assert.AreEqual(expectedCount, actualResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyElseBlock_HasWhitespace()
        {
            const string inputcode =
@"Sub Foo()
    If True Then
        Dim a as Long
        a = 0
    Else
    
    
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputcode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            const int expectedCount = 1;

            Assert.AreEqual(expectedCount, actualResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyElseBlock_HasDeclarationStatement()
        {
            const string inputCode =
@"Sub Foo()
    If True Then
        Dim a as Long
        a = 0
    Else
        Dim d
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            const int expectedCount = 1;

            Assert.AreEqual(expectedCount, actualResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyElseBlock_HasExecutableStatement()
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
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            const int expectedCount = 1;

            Assert.AreEqual(expectedCount, actualResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyElseBlock_QuickFixRemovesElse()
        {
            const string inputCode =
@"Sub Foo()
    If True Then
    Else
    End If
End Sub";

            const string expectedCode =
@"Sub Foo()
    If True Then
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveEmptyConditionBlockQuickFix(state).Fix(actualResults.First());
            var actualCode = state.GetRewriter(component).GetText();

            Assert.AreEqual(expectedCode, actualCode);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyElseBlock_QuickFixRemoveInLineIfThenElse()
        {
            const string inputCode =
@"Sub Foo()
    If True Then Else 
End Sub";

            const string expectedCode =
@"Sub Foo()
    If Not True Then 
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            var inspectionToFix = actualResults.First();
            new RemoveEmptyConditionBlockQuickFix(state).Fix(inspectionToFix);

            var actualCode = state.GetRewriter(component).GetText();
            Assert.AreEqual(expectedCode, actualCode);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void NonEmptyIfBlock_FindsNoResults()
        {
            const string inputcode =
@"Sub Foo()
    If True Then
        Dim a as String
        a = ""a""
    End If
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputcode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.IsFalse(actualResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void NonEmptyIAndElseifBlock_FindsNoResults()
        {
            const string inputcode =
@"Sub Foo()
    If True Then
        Dim a as String
        a = ""a""
    ElseIf True Then
        Dim b as String
        b = ""b""
    End If
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputcode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.IsFalse(actualResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void NonEmptyIfAndElseBlock_FindsNoResults()
        {
            const string inputcode =
@"Sub Foo()
    If True Then
        Dim a as String
        a = ""a""
    Else
        Dim b as String
        b = ""b""
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputcode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.IsFalse(actualResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void NonEmptyIfElseifElseBlock_FindsNoResults()
        {
            const string inputCode =
@"Sub Foo()
    If True Then
        Dim a as String
        a = ""a""
    ElseIf True Then
        Dim b as String
        b = ""b""
    Else
        Dim c as String
        c = ""c""
    End If
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.IsFalse(actualResults.Any());
        }
    }
}
