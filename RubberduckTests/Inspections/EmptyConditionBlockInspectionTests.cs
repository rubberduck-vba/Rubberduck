using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
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
            var inspection = new EmptyConditionBlockInspection(null, ConditionBlockToInspect.NA);
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = nameof(EmptyConditionBlockInspection);

            var inspection = new EmptyConditionBlockInspection(null, ConditionBlockToInspect.NA);

            Assert.AreEqual(inspectionName, inspection.Name);
        }

        #region EmptyIfBlock
        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_FiresOnEmptyIfBlock()
        {
            const string inputCode =
@"Sub Foo()
    If True Then
    End If
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state, ConditionBlockToInspect.If);
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var allInspection = new EmptyConditionBlockInspection(state, ConditionBlockToInspect.All);
            var allInspector = InspectionsHelper.GetInspector(allInspection);
            var allInspectionResults = allInspector.FindIssuesAsync(state, CancellationToken.None).Result;

            var elseifInspection = new EmptyConditionBlockInspection(state, ConditionBlockToInspect.ElseIf);
            var elseifInspector = InspectionsHelper.GetInspector(elseifInspection);
            var elseifInspectionResults = elseifInspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(3, allInspectionResults.Count());
            Assert.AreEqual(1, elseifInspectionResults.Count());
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state, ConditionBlockToInspect.Else);
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state, ConditionBlockToInspect.If);
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state, ConditionBlockToInspect.If);
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state, ConditionBlockToInspect.If);
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state,ConditionBlockToInspect.If);
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state, ConditionBlockToInspect.If);
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state, ConditionBlockToInspect.If);
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state, ConditionBlockToInspect.If);
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state, ConditionBlockToInspect.If);
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state,ConditionBlockToInspect.If);
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state,
                                    ConditionBlockToInspect.If & ConditionBlockToInspect.ElseIf);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.IsFalse(inspectionResults.Any());
        }
        #endregion

        #region EmptyElseBlock
        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyElseBlock_FiresOnEmptyIfBlock()
        {
            const string inputcode =
@"Sub Foo()
    If True Then
    End If
End Sub";
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputcode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state, ConditionBlockToInspect.If);
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputcode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state, ConditionBlockToInspect.All);
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
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state, ConditionBlockToInspect.Else);
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state, ConditionBlockToInspect.Else);
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state, ConditionBlockToInspect.Else);
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state,ConditionBlockToInspect.Else);
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputcode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state,ConditionBlockToInspect.Else);
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state, ConditionBlockToInspect.Else);
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state,ConditionBlockToInspect.Else);
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state,ConditionBlockToInspect.Else);
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyConditionBlockInspection(state,ConditionBlockToInspect.Else);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            var inspectionToFix = actualResults.First();
            new RemoveEmptyConditionBlockQuickFix(state).Fix(inspectionToFix);
            
            var actualCode = state.GetRewriter(component).GetText();
            Assert.AreEqual(expectedCode, actualCode);
        }
        #endregion
    }
}
