using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class EmptyIfBlockInspectionTests
    {
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

            var inspection = new EmptyIfBlockInspection(state);
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

            var inspection = new EmptyIfBlockInspection(state);
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
    Else
    End If
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyIfBlockInspection(state);
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

            var inspection = new EmptyIfBlockInspection(state);
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
    End If
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyIfBlockInspection(state);
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

            var inspection = new EmptyIfBlockInspection(state);
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

            var inspection = new EmptyIfBlockInspection(state);
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

            var inspection = new EmptyIfBlockInspection(state);
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
    End If
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyIfBlockInspection(state);
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

            var inspection = new EmptyIfBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_ElseIfBlockHasExecutableStatement()
        {
            const string inputCode =
@"Sub Foo()
    If True Then
        Dim d
    ElseIf False Then
        Dim b
    End If
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyIfBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_QuickFixRemovesLoneIf()
        {
            const string inputCode =
@"Sub Foo()
    If True Then
    End If
End Sub";
            const string expectedCode =
@"Sub Foo()
    
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyIfBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveEmptyIfBlockQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_QuickFixRemovesSingleLineIf()
        {
            const string inputCode =
@"Sub Foo()
    If True Then Else Bar
End Sub

Sub Bar()
End Sub";

            const string expectedCode =
@"Sub Foo()
    If Not (True) Then Bar
End Sub

Sub Bar()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyIfBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveEmptyIfBlockQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_QuickFixRemovesSingleLineIf_NoWhiteSpace()
        {
            const string inputCode =
@"Sub Foo()
    If True Then Else(Bar)
End Sub

Sub Bar()
End Sub";

            const string expectedCode =
@"Sub Foo()
    If Not (True) Then (Bar)
End Sub

Sub Bar()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyIfBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveEmptyIfBlockQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_QuickFixRemovesLoneIf_WithComment()
        {
            const string inputCode =
@"Sub Foo()
    If True Then
        ' Im a comment
    End If
End Sub";
            const string expectedCode =
@"Sub Foo()
    
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyIfBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveEmptyIfBlockQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_QuickFixRemovesIf_WithElseIfAndElse()
        {
            const string inputCode =
@"Sub Foo()
    If True Then
    ElseIf False Then
        Dim d
    Else
        Dim b
    End If
End Sub";
            const string expectedCode =
@"Sub Foo()
    If False Then
        Dim d
    Else
        Dim b
    End If
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyIfBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveEmptyIfBlockQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_QuickFixRemovesElseIf()
        {
            const string inputCode =
@"Sub Foo()
    If True Then
        Dim d
    ElseIf False Then
    Else
        Dim b
    End If
End Sub";
            const string expectedCode =
@"Sub Foo()
    If True Then
        Dim d
    Else
        Dim b
    End If
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyIfBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveEmptyIfBlockQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_QuickFixRemovesElseIf_HasComment()
        {
            const string inputCode =
@"Sub Foo()
    If True Then
        Dim d
    ElseIf False Then
        ' Im a comment
    Else
        Dim b
    End If
End Sub";
            const string expectedCode =
@"Sub Foo()
    If True Then
        Dim d
    Else
        Dim b
    End If
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyIfBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveEmptyIfBlockQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_QuickFixRemovesIf_UpdatesElseIf()
        {
            const string inputCode =
@"Sub Foo()
    If True Then
    ElseIf False Then
        Dim d
    End If
End Sub";
            const string expectedCode =
@"Sub Foo()
    If False Then
        Dim d
    End If
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyIfBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveEmptyIfBlockQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_QuickFixRemovesElse_InvertsIf_SimpleCondition()
        {
            const string inputCode =
@"Sub Foo()
    If True Then
    Else
    End If
End Sub";
            const string expectedCode =
@"Sub Foo()
    If Not (True) Then
    
    End If
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyIfBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveEmptyIfBlockQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_QuickFixRemovesElse_InvertsIf_Equals()
        {
            const string inputCode =
@"Sub Foo()
    If True = True Then
    Else
    End If
End Sub";
            const string expectedCode =
@"Sub Foo()
    If True <> True Then
    
    End If
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyIfBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveEmptyIfBlockQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_QuickFixRemovesElse_InvertsIf_NotEquals()
        {
            const string inputCode =
@"Sub Foo()
    If True <> True Then
    Else
    End If
End Sub";
            const string expectedCode =
@"Sub Foo()
    If True = True Then
    
    End If
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyIfBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveEmptyIfBlockQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_QuickFixRemovesElse_InvertsIf_LessThan()
        {
            const string inputCode =
@"Sub Foo()
    If True < True Then
    Else
    End If
End Sub";
            const string expectedCode =
@"Sub Foo()
    If True >= True Then
    
    End If
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyIfBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveEmptyIfBlockQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_QuickFixRemovesElse_InvertsIf_LessThanEquals()
        {
            const string inputCode =
@"Sub Foo()
    If True <= True Then
    Else
    End If
End Sub";
            const string expectedCode =
@"Sub Foo()
    If True > True Then
    
    End If
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyIfBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveEmptyIfBlockQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_QuickFixRemovesElse_InvertsIf_GreaterThan()
        {
            const string inputCode =
@"Sub Foo()
    If True > True Then
    Else
    End If
End Sub";
            const string expectedCode =
@"Sub Foo()
    If True <= True Then
    
    End If
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyIfBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveEmptyIfBlockQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_QuickFixRemovesElse_InvertsIf_GreaterThanEquals()
        {
            const string inputCode =
@"Sub Foo()
    If True >= True Then
    Else
    End If
End Sub";
            const string expectedCode =
@"Sub Foo()
    If True < True Then
    
    End If
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyIfBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveEmptyIfBlockQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_QuickFixRemovesElse_InvertsIf_Not()
        {
            const string inputCode =
@"Sub Foo()
    If Not True Then
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

            var inspection = new EmptyIfBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveEmptyIfBlockQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_QuickFixRemovesElse_InvertsIf_Not_NoWhitespace()
        {
            const string inputCode =
@"Sub Foo()
    If Not(True) Then
    Else
    End If
End Sub";
            const string expectedCode =
@"Sub Foo()
    If (True) Then
    
    End If
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyIfBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveEmptyIfBlockQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_QuickFixRemovesElse_InvertsIf_And()
        {
            const string inputCode =
@"Sub Foo()
    If True And True Then
    Else
    End If
End Sub";
            const string expectedCode =
@"Sub Foo()
    If True Or True Then
    
    End If
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyIfBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveEmptyIfBlockQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_QuickFixRemovesElse_InvertsIf_Or()
        {
            const string inputCode =
@"Sub Foo()
    If True Or True Then
    Else
    End If
End Sub";
            const string expectedCode =
@"Sub Foo()
    If True And True Then
    
    End If
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyIfBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveEmptyIfBlockQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_QuickFixRemovesElse_InvertsIf_Xor()
        {
            const string inputCode =
@"Sub Foo()
    If True Xor True Then
    Else
    End If
End Sub";
            const string expectedCode =
@"Sub Foo()
    If Not (True Xor True) Then
    
    End If
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyIfBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveEmptyIfBlockQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_QuickFixRemovesElse_InvertsIf_ComplexCondition()
        {
            const string inputCode =
@"Sub Foo()
    If True Or True And True Or True Then
    Else
    End If
End Sub";
            const string expectedCode =
@"Sub Foo()
    If True Or True And True And True Then
    
    End If
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyIfBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveEmptyIfBlockQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_QuickFixRemovesElse_InvertsIf_ComplexCondition1()
        {
            const string inputCode =
@"Sub Foo()
    If True And True Or True And True Then
    Else
    End If
End Sub";
            const string expectedCode =
@"Sub Foo()
    If True And True And True And True Then
    
    End If
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyIfBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveEmptyIfBlockQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_QuickFixRemovesElse_InvertsIf_ComplexCondition_WithParentheses()
        {
            const string inputCode =
@"Sub Foo()
    If (True Or True) And (True Or True) Then
    Else
    End If
End Sub";
            const string expectedCode =
@"Sub Foo()
    If (True Or True) Or (True Or True) Then
    
    End If
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyIfBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveEmptyIfBlockQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyIfBlock_QuickFixRemovesElse_InvertsIf_ComplexCondition2()
        {
            const string inputCode =
@"Sub Foo()
    If 1 > 2 And 3 = 3 Or 4 <> 5 And 8 - 6 = 2 Then
    Else
    End If
End Sub";
            const string expectedCode =
@"Sub Foo()
    If 1 > 2 And 3 = 3 And 4 <> 5 And 8 - 6 = 2 Then
    
    End If
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyIfBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveEmptyIfBlockQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new EmptyIfBlockInspection(null);
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = nameof(EmptyIfBlockInspection);
            var inspection = new EmptyIfBlockInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
