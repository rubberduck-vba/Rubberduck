using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections;
using Rubberduck.Inspections.QuickFixes;
using Moq;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Inspections.Resources;
using RubberduckTests.Mocks;
using System.Threading;
using Rubberduck.UI.Refactorings;
using System.Windows.Forms;

namespace RubberduckTests.Inspections
{
    // todo: streamline test cases, currently testing too many things at once, all tests break if newname strategy changes
    [TestClass]
    public class AssignedByValParameterMakeLocalCopyQuickFixTests
    {

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_LocalVariableAssignment()
        {
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As String)
    Let arg1 = ""test""
End Sub";

            const string expectedCode =
@"Public Sub Foo(ByVal arg1 As String)
Dim localArg1 As String
localArg1 = arg1
    Let localArg1 = ""test""
End Sub";

            var quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_LocalVariableAssignment_NameInUse()
        {
            //Punt if the user-defined or auto-generated name is already used in the method
            var inputCode =
@"
Public Sub Foo(ByVal arg1 As String)
    localArg1 = 6
    Let arg1 = ""test""
End Sub"
;

            var expectedCode =
@"
Public Sub Foo(ByVal arg1 As String)
Dim localArg12 As String
localArg12 = arg1
    localArg1 = 6
    Let localArg12 = ""test""
End Sub"
;

            var quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);

            //Punt if the user-defined or auto-generated name is already used in the method
            inputCode =
@"
Public Sub Foo(ByVal arg1 As String)
    Dim fooVar, localArg1 As Long
    Let arg1 = ""test""
End Sub"
;

            expectedCode =
@"
Public Sub Foo(ByVal arg1 As String)
Dim localArg12 As String
localArg12 = arg1
    Dim fooVar, localArg1 As Long
    Let localArg12 = ""test""
End Sub"
;

            quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);

            //handles line continuations
            inputCode =
@"
Public Sub Foo(ByVal arg1 As String)
    Dim fooVar, _
        localArg1 As Long
    Let arg1 = ""test""
End Sub"
            ;

            expectedCode =
@"
Public Sub Foo(ByVal arg1 As String)
Dim localArg12 As String
localArg12 = arg1
    Dim fooVar, _
        localArg1 As Long
    Let localArg12 = ""test""
End Sub"
;
            quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);

            //Punt if the user-defined or auto-generated name is already present as an parameter name
            var userEnteredName = "theSecondArg";

            inputCode =
@"
Public Sub Foo(ByVal arg1 As String, theSecondArg As Long)
    Let arg1 = ""test""
End Sub

Public Sub FooFight(ByRef arg1 As String)
    localArg1 = 6
    Let arg1 = ""test""
End Sub
"
;

            expectedCode =
@"
Public Sub Foo(ByVal arg1 As String, theSecondArg As Long)
    Let arg1 = ""test""
End Sub

Public Sub FooFight(ByRef arg1 As String)
    localArg1 = 6
    Let arg1 = ""test""
End Sub
"
;

            quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode, userEnteredName);
            Assert.AreEqual(expectedCode, quickFixResult);

            //Punt if the user-defined or auto-generated name is already present as an parameter name
            userEnteredName = "moduleScopeName";

            inputCode =
@"
Private moduleSopeName As Long


Public Sub Foo(ByVal arg1 As String, theSecondArg As Long)
    Let arg1 = ""test""
    moduleScopeName = arg1 & ""Foo""
End Sub

Public Sub FooFight(ByRef arg1 As String)
    localArg1 = 6
    Let arg1 = ""test""
End Sub
"
;

            expectedCode =
@"
Private moduleSopeName As Long


Public Sub Foo(ByVal arg1 As String, theSecondArg As Long)
    Let arg1 = ""test""
    moduleScopeName = arg1 & ""Foo""
End Sub

Public Sub FooFight(ByRef arg1 As String)
    localArg1 = 6
    Let arg1 = ""test""
End Sub
"
;

            quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode, userEnteredName);
            Assert.AreEqual(expectedCode, quickFixResult);

        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_LocalVariableAssignment_NameInUseOtherSub()
        {
            //Make sure the modified code stays within the specific method under repair
            const string inputCode =
@"
Public Function Bar2(ByVal arg2 As String) As String
    Dim arg1 As String
    Let arg1 = ""Test1""
    Bar2 = arg1
End Function

Public Sub Foo(ByVal arg1 As String)
    Let arg1 = ""test""
End Sub

Public Sub Bar(ByVal arg2 As String)
    Dim arg1 As String
    Let arg1 = ""Test2""
End Sub"
;

            const string expectedCode =
@"
Public Function Bar2(ByVal arg2 As String) As String
    Dim arg1 As String
    Let arg1 = ""Test1""
    Bar2 = arg1
End Function

Public Sub Foo(ByVal arg1 As String)
Dim localArg1 As String
localArg1 = arg1
    Let localArg1 = ""test""
End Sub

Public Sub Bar(ByVal arg2 As String)
    Dim arg1 As String
    Let arg1 = ""Test2""
End Sub"
;

            var quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_LocalVariableAssignment_NameInUseOtherProperty()
        {
            //Make sure the modified code stays within the specific method under repair
            var inputCode =
@"
Option Explicit
Private mBar as Long
Public Property Let Foo(ByVal bar As Long)
    bar = 42
    bar = bar * 2
    mBar = bar
End Property

Public Property Get Foo() As Long
    Dim bar as Long
    bar = 12
    Foo = mBar
End Property

Public Function bar() As Long
    bar = 42
End Function
";
            var expectedCode =
@"
Option Explicit
Private mBar as Long
Public Property Let Foo(ByVal bar As Long)
Dim localBar As Long
localBar = bar
    localBar = 42
    localBar = localBar * 2
    mBar = localBar
End Property

Public Property Get Foo() As Long
    Dim bar as Long
    bar = 12
    Foo = mBar
End Property

Public Function bar() As Long
    bar = 42
End Function
";

            var quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
            //variable name use checks do not 'leak' into adjacent procedure(s)
            inputCode =
@"
Public Sub Foo(ByVal arg1 As String)
    Let arg1 = ""test""
End Sub

Public Sub FooFight(ByRef arg1 As String)
    localArg1 = 6
    Let arg1 = ""test""
End Sub
"
;

            expectedCode =
@"
Public Sub Foo(ByVal arg1 As String)
Dim localArg1 As String
localArg1 = arg1
    Let localArg1 = ""test""
End Sub

Public Sub FooFight(ByRef arg1 As String)
    localArg1 = 6
    Let arg1 = ""test""
End Sub
"
;

            quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_LocalVariableAssignment_SimilarNamesIgnored()
        {
            //Make sure the modified code stays within the specific method under repair
            var inputCode =
@"
Option Explicit

Public Sub Foo(ByVal bar As Long)
    bar = 42
    bar = bar * 2
    Dim barBell as Long
    barBell = 6
    Dim isobar as Long
    isobar = 13
    Dim bar_candy as Long
    Dim candy_bar as Long
    Dim bar_after_bar as Long
    Dim barsAlot as string
    barsAlot = ""barsAlot:"" & CStr(isobar) & CStr(bar) & CStr(barBell)
    barsAlot = ""barsAlot:"" & CStr(isobar) & CStr( _
        bar) & CStr(barBell)
    total = bar + isobar + candy_bar + bar + bar_candy + barBell + _
            bar_after_bar + bar
bar = 7
    barsAlot = ""bar""
End Sub
";
            var expectedCode =
@"
Option Explicit

Public Sub Foo(ByVal bar As Long)
Dim localBar As Long
localBar = bar
    localBar = 42
    localBar = localBar * 2
    Dim barBell as Long
    barBell = 6
    Dim isobar as Long
    isobar = 13
    Dim bar_candy as Long
    Dim candy_bar as Long
    Dim bar_after_bar as Long
    Dim barsAlot as string
    barsAlot = ""barsAlot:"" & CStr(isobar) & CStr(localBar) & CStr(barBell)
    barsAlot = ""barsAlot:"" & CStr(isobar) & CStr( _
        localBar) & CStr(barBell)
    total = localBar + isobar + candy_bar + localBar + bar_candy + barBell + _
            bar_after_bar + localBar
localBar = 7
    barsAlot = ""bar""
End Sub
";

            var quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_ProperPlacementOfDeclaration()
        {

            var inputCode =
@"Sub DoSomething(_
    ByVal foo As Long, _
    ByRef _
        bar, _
    ByRef barbecue _
                    )
    foo = 4
    bar = barbecue * _
               bar + foo / barbecue
End Sub
";

            var expectedCode =
@"Sub DoSomething(_
    ByVal foo As Long, _
    ByRef _
        bar, _
    ByRef barbecue _
                    )
Dim localFoo As Long
localFoo = foo
    localFoo = 4
    bar = barbecue * _
               bar + localFoo / barbecue
End Sub
";
            var quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new AssignedByValParameterInspection(null,null);
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "AssignedByValParameterInspection";
            var inspection = new AssignedByValParameterInspection(null,null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }

        private string ApplyLocalVariableQuickFixToCodeFragment(string inputCode, string userEnteredName = "")
        {
            var vbe = BuildMockVBEStandardModuleForVBAFragment(inputCode);

            var mockDialogFactory = BuildMockDialogFactory(userEnteredName);

            var inspectionResults = GetInspectionResults(vbe.Object, mockDialogFactory.Object);
            var result = inspectionResults.FirstOrDefault();
            if (result == null)
            {
                Assert.Inconclusive("Inspection yielded no results.");
            }

            result.QuickFixes.Single(s => s is AssignedByValParameterMakeLocalCopyQuickFix).Fix();

            return GetModuleContent(vbe.Object);
        }

        private Mock<IVBE> BuildMockVBEStandardModuleForVBAFragment(string inputCode)
        {
            var builder = new MockVbeBuilder();
            IVBComponent component;
            return builder.BuildFromSingleStandardModule(inputCode, out component);
        }

        private IEnumerable<Rubberduck.Inspections.Abstract.InspectionResultBase> GetInspectionResults(IVBE vbe, IAssignedByValParameterQuickFixDialogFactory mockDialogFactory)
        {
            var parser = MockParser.Create(vbe, new RubberduckParserState(vbe));
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new AssignedByValParameterInspection(parser.State, mockDialogFactory);
            return inspection.GetInspectionResults();
        }

        private Mock<IAssignedByValParameterQuickFixDialogFactory> BuildMockDialogFactory(string userEnteredName)
        {
            var mockDialog = new Mock<IAssignedByValParameterQuickFixDialog>();

            mockDialog.SetupAllProperties();

            if (userEnteredName.Length > 0)
            {
                mockDialog.SetupGet(m => m.NewName).Returns(() => userEnteredName);
            }
            mockDialog.SetupGet(m => m.DialogResult).Returns(() => DialogResult.OK);

            var mockDialogFactory = new Mock<IAssignedByValParameterQuickFixDialogFactory>();
            mockDialogFactory.Setup(f => f.Create(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<IEnumerable<string>>())).Returns(mockDialog.Object);

            return mockDialogFactory;
        }

        private string GetModuleContent(IVBE vbe)
        {
            var project = vbe.VBProjects[0];
            var module = project.VBComponents[0].CodeModule;
            return module.Content();
        }
    }
}
