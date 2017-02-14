using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Inspections;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using System.Collections.Generic;
using System;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class AssignedByValParameterInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_ReturnsResult_Sub()
        {
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As String)
    Let arg1 = ""test""
End Sub";

            AssertVbaFragmentYieldsExpectedInspectionResultCount(inputCode, 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_ReturnsResult_Function()
        {
            const string inputCode =
@"Function Foo(ByVal arg1 As Integer) As Boolean
    Let arg1 = 9
End Function";

            AssertVbaFragmentYieldsExpectedInspectionResultCount(inputCode, 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_ReturnsResult_MultipleParams()
        {
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As String, ByVal arg2 As Integer)
    Let arg1 = ""test""
    Let arg2 = 9
End Sub";

            AssertVbaFragmentYieldsExpectedInspectionResultCount(inputCode, 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_DoesNotReturnResult()
        {
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As String)
End Sub";

            AssertVbaFragmentYieldsExpectedInspectionResultCount(inputCode, 0);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_Ignored_DoesNotReturnResult_Sub()
        {
            const string inputCode =
@"'@Ignore AssignedByValParameter
Public Sub Foo(ByVal arg1 As String)
    Let arg1 = ""test""
End Sub";

            AssertVbaFragmentYieldsExpectedInspectionResultCount(inputCode, 0);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_ReturnsResult_SomeAssignedByValParams()
        {
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As String, ByVal arg2 As Integer)
    Let arg1 = ""test""
    
    Dim var1 As Integer
    var1 = arg2
End Sub";

            AssertVbaFragmentYieldsExpectedInspectionResultCount(inputCode, 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_QuickFixWorks()
        {
            string inputCode =
@"Public Sub Foo(ByVal barByVal As String)
    Let barByVal = ""test""
End Sub";
            string expectedCode =
@"Public Sub Foo(ByRef barByVal As String)
    Let barByVal = ""test""
End Sub";

            var quickFixResult = ApplyPassParameterByReferenceQuickFixToVBAFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);

            //check when ByVal argument is one of several parameters
            inputCode =
@"Public Sub Foo(ByRef firstArg As Long, ByVal barByVal As String, secondArg as Double)
    Let barByVal = ""test""
End Sub";
            expectedCode =
@"Public Sub Foo(ByRef firstArg As Long, ByRef barByVal As String, secondArg as Double)
    Let barByVal = ""test""
End Sub";

            quickFixResult = ApplyPassParameterByReferenceQuickFixToVBAFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_IgnoreQuickFixWorks()
        {
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As String)
    Let arg1 = ""test""
End Sub";

            const string expectedCode =
@"'@Ignore AssignedByValParameter
Public Sub Foo(ByVal arg1 As String)
    Let arg1 = ""test""
End Sub";

            var quickFixResult = ApplyIgnoreOnceQuickFixToVBAFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult); 
        }

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
            const string inputCode =
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
            const string expectedCode =
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
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_LocalVariableAssignment_SimilarNamesIgnored()
        {
            //Make sure the modified code stays within the specific method under repair
            const string inputCode =
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
            const string expectedCode =
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
            var inspection = new AssignedByValParameterInspection(null);
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "AssignedByValParameterInspection";
            var inspection = new AssignedByValParameterInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }

        private string ApplyPassParameterByReferenceQuickFixToVBAFragment(string inputCode)
        {
            var vbe = BuildMockVBEStandardModuleForVBAFragment(inputCode);
            var inspectionResults = GetInspectionResults(vbe);

            inspectionResults.First().QuickFixes.Single(s => s is PassParameterByReferenceQuickFix).Fix();

            return GetModuleContent(vbe);
        }

        private string ApplyLocalVariableQuickFixToVBAFragment(string inputCode)
        {
            var vbe = BuildMockVBEStandardModuleForVBAFragment(inputCode);
            var inspectionResults = GetInspectionResults(vbe);

            var quickFixBase = inspectionResults.First().QuickFixes.Single(s => s is AssignedByValParameterQuickFix);
            AssignedByValParameterQuickFix assignByValParamQFix = (AssignedByValParameterQuickFix)quickFixBase;

            assignByValParamQFix.TESTONLY_FixUsingAutoGeneratedName();

            return GetModuleContent(vbe);
        }
        private string ApplyIgnoreOnceQuickFixToVBAFragment(string inputCode)
        {
            var vbe = BuildMockVBEStandardModuleForVBAFragment(inputCode);
            var inspectionResults = GetInspectionResults(vbe);

            inspectionResults.First().QuickFixes.Single(s => s is IgnoreOnceQuickFix).Fix();

            return GetModuleContent(vbe);
        }
        private string GetModuleContent(Mock<IVBE> vbe)
        {
            var project = vbe.Object.VBProjects[0];
            var module = project.VBComponents[0].CodeModule;
            return module.Content();
        }
        private IEnumerable<Rubberduck.Inspections.Abstract.InspectionResultBase> GetInspectionResults(string inputCode)
        {
            var vbe = BuildMockVBEStandardModuleForVBAFragment(inputCode);
            return GetInspectionResults(vbe);
        }
        private IEnumerable<Rubberduck.Inspections.Abstract.InspectionResultBase> GetInspectionResults(Mock<IVBE> vbe)
        {
            var parser = GetMockParseCoordinator(vbe);
            var inspection = new AssignedByValParameterInspection(parser.State);
            return inspection.GetInspectionResults();
        }
        private void AssertVbaFragmentYieldsExpectedInspectionResultCount(string inputCode, int expectedCount)
        {
            var inspectionResults = GetInspectionResults(inputCode);
            Assert.AreEqual(expectedCount, inspectionResults.Count());
        }
        private Mock<IVBE> BuildMockVBEStandardModuleForVBAFragment(string inputCode)
        {
            var builder = new MockVbeBuilder();
            IVBComponent component;
            return builder.BuildFromSingleStandardModule(inputCode, out component);

        }
        private ParseCoordinator GetMockParseCoordinator(Mock<IVBE> vbe)
        {
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }
            return parser;
        }
    }
}
