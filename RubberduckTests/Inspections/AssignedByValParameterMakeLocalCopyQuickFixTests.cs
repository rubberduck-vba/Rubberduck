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


namespace RubberduckTests.Inspections
{

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
Dim xArg1 As String
xArg1 = arg1
    Let xArg1 = ""test""
End Sub";

            var quickFixResult = ApplyLocalVariableQuickFixToVBAFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_LocalVariableAssignment_NameInUse()
        {
            //Punt if the user-defined or auto-generated name is already used in the method
            string inputCode =
@"
Public Sub Foo(ByVal arg1 As String)
    xArg1 = 6
    Let arg1 = ""test""
End Sub"
;

            string expectedCode =
@"
Public Sub Foo(ByVal arg1 As String)
    xArg1 = 6
    Let arg1 = ""test""
End Sub"
;

            var quickFixResult = ApplyLocalVariableQuickFixToVBAFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);

            //Punt if the user-defined or auto-generated name is already used in the method
            inputCode =
@"
Public Sub Foo(ByVal arg1 As String)
    Dim fooVar, xArg1 As Long
    Let arg1 = ""test""
End Sub"
;

            expectedCode =
@"
Public Sub Foo(ByVal arg1 As String)
    Dim fooVar, xArg1 As Long
    Let arg1 = ""test""
End Sub"
;

            quickFixResult = ApplyLocalVariableQuickFixToVBAFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);

            //handles line continuations
            inputCode =
@"
Public Sub Foo(ByVal arg1 As String)
    Dim fooVar, _
        xArg1 As Long
    Let arg1 = ""test""
End Sub"
            ;

            expectedCode =
@"
Public Sub Foo(ByVal arg1 As String)
    Dim fooVar, _
        xArg1 As Long
    Let arg1 = ""test""
End Sub"
;
            quickFixResult = ApplyLocalVariableQuickFixToVBAFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);

            //Punt if the user-defined or auto-generated name is already present as an parameter name
            string userEnteredName = "theSecondArg";

            inputCode =
@"
Public Sub Foo(ByVal arg1 As String, theSecondArg As Long)
    Let arg1 = ""test""
End Sub

Public Sub FooFight(ByRef arg1 As String)
    xArg1 = 6
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
    xArg1 = 6
    Let arg1 = ""test""
End Sub
"
;

            quickFixResult = ApplyLocalVariableQuickFixToVBAFragment(inputCode, userEnteredName);
            Assert.AreEqual(expectedCode, quickFixResult);

            //Punt if the user-defined or auto-generated name is already present as an parameter name
            userEnteredName = "theSecondArg";

            inputCode =
@"
Private moduleSopeName As Long


Public Sub Foo(ByVal arg1 As String, theSecondArg As Long)
    Let arg1 = ""test""
    moduleScopeName = arg1 & ""Foo""
End Sub

Public Sub FooFight(ByRef arg1 As String)
    xArg1 = 6
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
    xArg1 = 6
    Let arg1 = ""test""
End Sub
"
;

            quickFixResult = ApplyLocalVariableQuickFixToVBAFragment(inputCode, userEnteredName);
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
Dim xArg1 As String
xArg1 = arg1
    Let xArg1 = ""test""
End Sub

Public Sub Bar(ByVal arg2 As String)
    Dim arg1 As String
    Let arg1 = ""Test2""
End Sub"
;

            var quickFixResult = ApplyLocalVariableQuickFixToVBAFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_LocalVariableAssignment_NameInUseOtherProperty()
        {
            //Make sure the modified code stays within the specific method under repair
            string inputCode =
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
            string expectedCode =
@"
Option Explicit
Private mBar as Long
Public Property Let Foo(ByVal bar As Long)
Dim xBar As Long
xBar = bar
    xBar = 42
    xBar = xBar * 2
    mBar = xBar
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

            var quickFixResult = ApplyLocalVariableQuickFixToVBAFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
            //variable name use checks do not 'leak' into adjacent procedure(s)
            inputCode =
@"
Public Sub Foo(ByVal arg1 As String)
    Let arg1 = ""test""
End Sub

Public Sub FooFight(ByRef arg1 As String)
    xArg1 = 6
    Let arg1 = ""test""
End Sub
"
;

            expectedCode =
@"
Public Sub Foo(ByVal arg1 As String)
Dim xArg1 As String
xArg1 = arg1
    Let xArg1 = ""test""
End Sub

Public Sub FooFight(ByRef arg1 As String)
    xArg1 = 6
    Let arg1 = ""test""
End Sub
"
;

            quickFixResult = ApplyLocalVariableQuickFixToVBAFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);

        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_LocalVariableAssignment_SimilarNamesIgnored()
        {
            //Make sure the modified code stays within the specific method under repair
            string inputCode =
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
            string expectedCode =
@"
Option Explicit

Public Sub Foo(ByVal bar As Long)
Dim xBar As Long
xBar = bar
    xBar = 42
    xBar = xBar * 2
    Dim barBell as Long
    barBell = 6
    Dim isobar as Long
    isobar = 13
    Dim bar_candy as Long
    Dim candy_bar as Long
    Dim bar_after_bar as Long
    Dim barsAlot as string
    barsAlot = ""barsAlot:"" & CStr(isobar) & CStr(xBar) & CStr(barBell)
    barsAlot = ""barsAlot:"" & CStr(isobar) & CStr( _
        xBar) & CStr(barBell)
    total = xBar + isobar + candy_bar + xBar + bar_candy + barBell + _
            bar_after_bar + xBar
xBar = 7
    barsAlot = ""bar""
End Sub
";

            var quickFixResult = ApplyLocalVariableQuickFixToVBAFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
        }
        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_ProperPlacementOfDeclaration()
        {

            string inputCode =
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

            string expectedCode =
@"Sub DoSomething(_
    ByVal foo As Long, _
    ByRef _
        bar, _
    ByRef barbecue _
                    )
Dim xFoo As Long
xFoo = foo
    xFoo = 4
    bar = barbecue * _
               bar + xFoo / barbecue
End Sub
";
            var quickFixResult = ApplyLocalVariableQuickFixToVBAFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_LocalVariableAssignment_FunctionReturn()
        {
            const string inputCode =
@"Private Function MessingWithByValParameters(leaveAlone As Integer, ByVal messWithThis As String) As String
    If leaveAlone > 10 Then
        messWithThis = messWithThis & CStr(leaveAlone)
        messWithThis = Replace(messWithThis, ""OK"", ""yes"")
    End If
    MessingWithByValParameters = messWithThis
End Function
";

            const string expectedCode =
@"Private Function MessingWithByValParameters(leaveAlone As Integer, ByVal messWithThis As String) As String
Dim xMessWithThis As String
xMessWithThis = messWithThis
    If leaveAlone > 10 Then
        xMessWithThis = xMessWithThis & CStr(leaveAlone)
        xMessWithThis = Replace(xMessWithThis, ""OK"", ""yes"")
    End If
    MessingWithByValParameters = xMessWithThis
End Function
";
            var quickFixResult = ApplyLocalVariableQuickFixToVBAFragment(inputCode);
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

        private string ApplyLocalVariableQuickFixToVBAFragment(string inputCode, string userEnteredName = "")
        {
            var vbe = BuildMockVBEStandardModuleForVBAFragment(inputCode);
            var inspectionResults = GetInspectionResults(vbe, userEnteredName);

            inspectionResults.First().QuickFixes.Single(s => s is AssignedByValParameterMakeLocalCopyQuickFix).Fix();

            return GetModuleContent(vbe);
        }

        private Mock<IVBE> BuildMockVBEStandardModuleForVBAFragment(string inputCode)
        {
            var builder = new MockVbeBuilder();
            IVBComponent component;
            return builder.BuildFromSingleStandardModule(inputCode, out component);
        }
        private IEnumerable<Rubberduck.Inspections.Abstract.InspectionResultBase> GetInspectionResults(Mock<IVBE> vbe, string userEnteredName)
        {
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new AssignedByValParameterInspection(parser.State,new AssignedByValParameterQuickFixMockDialogFactory(userEnteredName));
            return inspection.GetInspectionResults();
        }

        private string GetModuleContent(Mock<IVBE> vbe)
        {
            var project = vbe.Object.VBProjects[0];
            var module = project.VBComponents[0].CodeModule;
            return module.Content();
        }
    }
}
