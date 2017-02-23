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
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers;

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
        public void AssignedByValParameter_NoResultForLeftHandSideMemberAssignment()
        {
            var class1 = @"
Option Explicit
Private mSomething As Long
Public Property Get Something() As Long
    Something = mSomething
End Property
Public Property Let Something(ByVal value As Long)
    mSomething = value
End Property
";
            var caller = @"
Option Explicit
Private Sub DoSomething(ByVal foo As Class1)
    foo.Something = 42
End Sub
";
            var builder = new MockVbeBuilder();
            var vbe = builder.ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Module1", ComponentType.StandardModule, caller)
                .MockVbeBuilder()
                .Build();
            var results = GetInspectionResults(vbe);
            Assert.AreEqual(0, results.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_QuickFixWorks()
        {

            string inputCode =
@"Public Sub Foo(Optional ByVal barByVal As String = ""XYZ"")
    Let barByVal = ""test""
End Sub";
            string expectedCode =
@"Public Sub Foo(Optional ByRef barByVal As String = ""XYZ"")
    Let barByVal = ""test""
End Sub";

            var quickFixResult = ApplyPassParameterByReferenceQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);

            //check when ByVal argument is one of several parameters
            inputCode =
@"Public Sub Foo(ByRef firstArg As Long, Optional ByVal barByVal As String = """", secondArg as Double)
    Let barByVal = ""test""
End Sub";
            expectedCode =
@"Public Sub Foo(ByRef firstArg As Long, Optional ByRef barByVal As String = """", secondArg as Double)
    Let barByVal = ""test""
End Sub";

            quickFixResult = ApplyPassParameterByReferenceQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);

            inputCode =
@"
Private Sub Foo(Optional ByVal  _
    bar _
    As _
    Long = 4, _
    ByVal _
    barTwo _
    As _
    Long)
bar = 42
End Sub
"
;
            expectedCode =
@"
Private Sub Foo(Optional ByRef  _
    bar _
    As _
    Long = 4, _
    ByVal _
    barTwo _
    As _
    Long)
bar = 42
End Sub
"
;
            quickFixResult = ApplyPassParameterByReferenceQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);

            inputCode =
@"Private Sub Foo(ByVal barByVal As Long, ByVal _xByValbar As Long,  ByVal _
    barTwo _
    As _
    Long)
barTwo = 42
End Sub
";
            expectedCode =
@"Private Sub Foo(ByVal barByVal As Long, ByVal _xByValbar As Long,  ByRef _
    barTwo _
    As _
    Long)
barTwo = 42
End Sub
";
            
            quickFixResult = ApplyPassParameterByReferenceQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);

            inputCode =
@"Private Sub Foo(ByVal barByVal As Long, ByVal barTwoon As Long,  ByVal _
    barTwo _
    As _
    Long)
barTwo = 42
End Sub
";
            expectedCode =
@"Private Sub Foo(ByVal barByVal As Long, ByVal barTwoon As Long,  ByRef _
    barTwo _
    As _
    Long)
barTwo = 42
End Sub
";

            quickFixResult = ApplyPassParameterByReferenceQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);

            inputCode =
@"Private Sub Foo(ByVal barByVal As Long, ByVal barTwoon As Long,  ByVal barTwo _
    As _
    Long)
barTwo = 42
End Sub
";
            expectedCode =
@"Private Sub Foo(ByVal barByVal As Long, ByVal barTwoon As Long,  ByRef barTwo _
    As _
    Long)
barTwo = 42
End Sub
";

            quickFixResult = ApplyPassParameterByReferenceQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);

            inputCode =
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

            expectedCode =
@"Sub DoSomething(_
    ByRef foo As Long, _
    ByRef _
        bar, _
    ByRef barbecue _
                    )
    foo = 4
    bar = barbecue * _
               bar + foo / barbecue
End Sub
";
            quickFixResult = ApplyPassParameterByReferenceQuickFixToCodeFragment(inputCode);
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

            var quickFixResult = ApplyIgnoreOnceQuickFixToCodeFragment(inputCode);
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


        private string ApplyPassParameterByReferenceQuickFixToCodeFragment(string inputCode)
        {
            var vbe = BuildMockVBEStandardModuleForVBAFragment(inputCode);
            var inspectionResults = GetInspectionResults(vbe);

            inspectionResults.First().QuickFixes.Single(s => s is PassParameterByReferenceQuickFix).Fix();

            return GetModuleContent(vbe);
        }


        private string ApplyIgnoreOnceQuickFixToCodeFragment(string inputCode)
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
