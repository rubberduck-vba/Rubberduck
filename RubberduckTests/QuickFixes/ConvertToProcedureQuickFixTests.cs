using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    [TestClass]
    public class ConvertToProcedureQuickFixTests
    {
        [TestMethod]
        [TestCategory("QuickFixes")]
        [TestCategory("Unused Value")]
        public void FunctionReturnValueNotUsed_QuickFixWorks_NoInterface()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Boolean
    If True Then
        Foo = _
        True
    Else
        Foo = False
    End If
End Function";

            const string expectedCode =
@"Public Sub Foo(ByVal bar As String)
    If True Then
        
    Else
        
    End If
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new FunctionReturnValueNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            new ConvertToProcedureQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("QuickFixes")]
        [TestCategory("Unused Value")]
        public void FunctionReturnValueNotUsed_QuickFixWorks_NoInterface_ManyBodyStatements()
        {
            const string inputCode =
@"Function foo(ByRef fizz As Boolean) As Boolean
    fizz = True
    goo
label1:
    foo = fizz
End Function

Sub goo()
End Sub";

            const string expectedCode =
@"Sub foo(ByRef fizz As Boolean)
    fizz = True
    goo
label1:
    
End Sub

Sub goo()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new FunctionReturnValueNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            new ConvertToProcedureQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("QuickFixes")]
        [TestCategory("Unused Value")]
        public void FunctionReturnValueNotUsed_QuickFixWorks_Interface()
        {
            const string inputInterfaceCode =
@"Public Function Test() As Integer
End Function";

            const string expectedInterfaceCode =
@"Public Sub Test()
End Sub";

            const string inputImplementationCode1 =
@"Implements IFoo
Public Function IFoo_Test() As Integer
    IFoo_Test = 42
End Function";

            const string inputImplementationCode2 =
@"Implements IFoo
Public Function IFoo_Test() As Integer
    IFoo_Test = 42
End Function";

            const string callSiteCode =
@"
Public Function Baz()
    Dim testObj As IFoo
    Set testObj = new Bar
    testObj.Test
End Function";

            var builder = new MockVbeBuilder();
            var vbe = builder.ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                             .AddComponent("IFoo", ComponentType.ClassModule, inputInterfaceCode)
                             .AddComponent("Bar", ComponentType.ClassModule, inputImplementationCode1)
                             .AddComponent("Bar2", ComponentType.ClassModule, inputImplementationCode2)
                             .AddComponent("TestModule", ComponentType.StandardModule, callSiteCode)
                             .MockVbeBuilder().Build();

            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new FunctionReturnValueNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            new ConvertToProcedureQuickFix(state).Fix(inspectionResults.First());

            var component = vbe.Object.VBProjects[0].VBComponents[0];
            Assert.AreEqual(expectedInterfaceCode, state.GetRewriter(component).GetText());
        }
    }
}
